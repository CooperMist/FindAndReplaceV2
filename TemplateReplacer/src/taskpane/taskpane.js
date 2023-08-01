/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
	if (info.host === Office.HostType.Word) {
		document.getElementById("sideload-msg").style.display = "none";
		document.getElementById("app-body").style.display = "flex";
		document.getElementById("add").onclick = () => addRowToTable("", "");
		document.getElementById("remove-selected").onclick = () => removeSelectedTags();
		document.getElementById("replace-all").onclick = () => replaceDocumentTags(false);
		document.getElementById("replace-selected").onclick = () => replaceDocumentTags(true);		
		document.getElementById("autopopulate").onclick = autopopulateTags;
		
		// document.getElementById("add").onclick = () => addRowToTable("", "");
		// loadSampleData("This is a sample text [[inserted]] in [[the]] document");
	}	
});


export function loadSampleData(text) {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {
        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a commmand to clear the contents of the body.
        body.clear();
        // Queue a command to insert text into the end of the Word document body.
        body.insertText(text, Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
        return context.sync();
    })
}

export function addToTagTable(tags) {
  console.log(tags);
  tags.forEach(function (tag) {
      addRowToTable(tag, "");
  });
}

export function addRowToTable(leftValue, rightValue) {
  var table = document.getElementById("tag-table");  // According to Home.html, that's the table id.
  let row = document.createElement('tr');
  var columns = Array.of();
  
  let checkColumn = document.createElement('td');
  var selected = document.createElement("input");
  selected.type = "checkbox";
  selected.checked = false;
  checkColumn.append(selected);
  row.append(checkColumn);
  columns.push(checkColumn);

  for (let i = 0; i < 2; i++) {
      let col = document.createElement('td');
      let field = document.createElement("input");
      field.type = "text";
      field.value = i != 0 ? rightValue : leftValue;
      col.append(field);
      row.append(col);
      columns.push(col);
  }
  table.append(row);
  return columns;
}

export function getFirstSelectedTag(rows) {
	for (let i = 0; i < rows.length; i++) {
		var contents = getRowContents(rows[i]);
		if (contents['selected'])
			return i;
	}
	return -1;
	
}

export function removeSelectedTags() {
	var table = document.getElementById("tag-table");
	var index = getFirstSelectedTag(table.rows);
	while (index != -1) {
		table.deleteRow(index);
		index = getFirstSelectedTag(table.rows);
	}
}

export function replaceInDocument(original, replaced) {
	original = "[[" + original + "]]";
	console.log("Replacing " + original + " with " + replaced);
	Word.run(async function (context) {
		var options = Word.SearchOptions.newObject(context);
		options.matchCase = false;
		var tagResults = context.document.body.search(original, options);
		context.load(tagResults, 'text');
		return context.sync()
		.then(() => {
			console.log("Found " + original + " " + tagResults.items.length + " times.");
			for (let v = 0; v < tagResults.items.length; v++) {
				tagResults.items[v].insertText(replaced, Word.InsertLocation.replace);
			}

		}).then(context.sync);
  });
}

export function getTableContents(ignoreUnselected) {
	var table = document.getElementById("tag-table");  // According to Home.html, that's the table id.
	var currentRowsInTable = table.rows.length;  // max amount of rows.

	var tags = {};

  // this loop populates the tags/values in a K-V dict system.
  	for (let i = 1; i < currentRowsInTable; i++) {
    	var contents = getRowContents(table.rows[i]);
    	if (ignoreUnselected && !contents['selected'])
        	continue;
      	tags[contents['left']] = contents['right'];
     	// tags[table.rows[i].cells[1].children[0].value] = table.rows[i].cells[2].children[0].value;
  }
  return tags;
}

export function findTags(text, excludeList, leftMarker, rightMarker, leftLength, rightLength) {
	var startPosition = 0;
	var tags = Array.of();
	while (true) {
		// find the next tag
		// var next = text.substring(startPosition).search(leftMarker + "[a-zA-Z]+" + rightMarker);
		var next = text.substring(startPosition).search(leftMarker + "[a-zA-Z]+" + rightMarker);
		if (next == -1)
			break;
		startPosition += next;
		// get the tag name from inside the markers.  
		var leftPosition = text.substring(startPosition).search(/\[\[/g);
		var tag = text.substring(startPosition + leftLength + leftPosition, startPosition + leftLength + text.substring(startPosition + leftLength).search(rightMarker));
		if (!tags.includes(tag) && !excludeList.includes(tag))
			tags.push(tag);
		startPosition += leftLength + tag.length + rightLength;
	}
	return tags;
	}

export function getRowContents(row) {
  return {
    'selected': row.cells[0].children[0].checked,
    'left': row.cells[1].children[0].value,
    'right': row.cells[2].children[0].value
  };
}

export function autopopulateTags() {
  	Word.run(async function (context) {
      	var documentBody = context.document.body;
		context.load(documentBody);
		// Synchronize the document state by executing the queued commands
		// and return a promise to indicate task completion.
		return context.sync()
			.then(() => {

				var t = documentBody.text;
				// get all the keys not in the table already
				var tags = findTags(documentBody.text, Object.keys(getTableContents()), "[[", "]]", 2, 2);
				addToTagTable(tags);
			});
			
		});
}

export async function replaceDocumentTags(ignoreUnselected) {
	var tags = getTableContents(ignoreUnselected);
	for (const [key, value] of Object.entries(tags)) {
		replaceInDocument(key, value);
	}
}