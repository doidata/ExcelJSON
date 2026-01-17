/**
 * ExcelJSON for Google Sheets
 * Converts structured spreadsheet data to d3.js-compatible treemap JSON
 *
 * This is a Google Apps Script port of the VBA ExcelJSON macro
 */

/**
 * Adds a custom menu to Google Sheets when the spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ExcelJSON')
      .addItem('Generate JSON', 'generateJSON')
      .addItem('Show JSON in Sidebar', 'showJSONInSidebar')
      .addToUi();
}

/**
 * Main function to generate JSON from spreadsheet data
 * Reads from Sheet 1 (nodes) and Sheet 2 (children) and generates d3.js treemap JSON
 */
function generateJSON() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheets()[0]; // First sheet (nodes)
    var sheet2 = ss.getSheets()[1]; // Second sheet (children)

    // Build the JSON structure
    var json = buildTreeJSON(sheet1, sheet2);

    // Output to a new sheet or update existing output sheet
    outputJSON(ss, json);

    SpreadsheetApp.getUi().alert('Success!', 'JSON generated successfully. Check the "JSON Output" sheet.', SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error processing data: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error: ' + error.message + '\n' + error.stack);
  }
}

/**
 * Shows the generated JSON in a sidebar
 */
function showJSONInSidebar() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheets()[0];
    var sheet2 = ss.getSheets()[1];

    var json = buildTreeJSON(sheet1, sheet2);
    var jsonString = JSON.stringify(json, null, 2);

    var html = HtmlService.createHtmlOutput('<pre style="font-family: monospace; font-size: 12px; white-space: pre-wrap; word-wrap: break-word;">' +
                                           escapeHtml(jsonString) +
                                           '</pre>')
        .setTitle('d3.js Treemap JSON')
        .setWidth(400);

    SpreadsheetApp.getUi().showSidebar(html);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error processing data: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error: ' + error.message + '\n' + error.stack);
  }
}

/**
 * Builds the tree JSON structure from the two sheets
 * @param {Sheet} sheet1 - The nodes sheet
 * @param {Sheet} sheet2 - The children sheet
 * @return {Object} The root node of the tree
 */
function buildTreeJSON(sheet1, sheet2) {
  var nodes = {};

  // Read nodes from Sheet 1
  var data1 = sheet1.getDataRange().getValues();
  var headers1 = data1[0];

  for (var i = 1; i < data1.length; i++) {
    var row = data1[i];
    if (!row[0]) break; // Stop at first empty row

    var nodeId = row[0].toString();
    var parentId = row[1] ? row[1].toString() : 'root';

    // Create node object
    var node = {
      id: nodeId,
      parent: parentId,
      name: nodeId,
      children: [],
      attributes: {}
    };

    // Add attributes from columns 3 onwards
    for (var j = 2; j < headers1.length && headers1[j]; j++) {
      if (row[j] !== null && row[j] !== '') {
        node.attributes[headers1[j]] = row[j];
      }
    }

    nodes[nodeId] = node;
  }

  // Build parent-child relationships
  for (var nodeId in nodes) {
    var node = nodes[nodeId];
    if (node.parent && node.parent !== 'root' && nodes[node.parent]) {
      nodes[node.parent].children.push(node);
    }
  }

  // Read children data from Sheet 2 (if exists)
  if (sheet2) {
    try {
      var data2 = sheet2.getDataRange().getValues();
      var headers2 = data2[0];

      for (var i = 1; i < data2.length; i++) {
        var row = data2[i];
        if (!row[0]) break; // Stop at first empty row

        var parentId = row[0].toString();

        if (nodes[parentId]) {
          // Create child object
          var child = {
            parent: parentId,
            attributes: {}
          };

          // Add attributes from columns 2 onwards
          for (var j = 1; j < headers2.length && headers2[j]; j++) {
            if (row[j] !== null && row[j] !== '') {
              child.attributes[headers2[j]] = row[j];
            }
          }

          nodes[parentId].children.push(child);
        }
      }
    } catch (e) {
      Logger.log('Sheet 2 processing skipped: ' + e.message);
    }
  }

  // Find and format root node(s)
  var rootNodes = [];
  for (var nodeId in nodes) {
    if (nodes[nodeId].parent === 'root') {
      rootNodes.push(formatNodeForD3(nodes[nodeId]));
    }
  }

  // Return the first root node, or an error if none found
  if (rootNodes.length === 0) {
    throw new Error('No root node found. Make sure at least one node has parent = "root"');
  }

  return rootNodes[0];
}

/**
 * Formats a node for d3.js treemap
 * @param {Object} node - The node to format
 * @return {Object} The formatted node
 */
function formatNodeForD3(node) {
  var result = {
    name: node.name
  };

  // Add attributes to the node
  for (var key in node.attributes) {
    result[key] = node.attributes[key];
  }

  // Add children
  if (node.children && node.children.length > 0) {
    result.children = node.children.map(function(child) {
      if (child.id) {
        // It's a node, recurse
        return formatNodeForD3(child);
      } else {
        // It's a child data object, just return the attributes
        var childObj = {};
        for (var key in child.attributes) {
          childObj[key] = child.attributes[key];
        }
        return childObj;
      }
    });
  }

  return result;
}

/**
 * Outputs the JSON to a dedicated sheet
 * @param {Spreadsheet} ss - The active spreadsheet
 * @param {Object} json - The JSON object to output
 */
function outputJSON(ss, json) {
  var jsonString = JSON.stringify(json, null, 2);

  // Get or create output sheet
  var outputSheet = ss.getSheetByName('JSON Output');
  if (!outputSheet) {
    outputSheet = ss.insertSheet('JSON Output');
  }

  // Clear existing content
  outputSheet.clear();

  // Write JSON to the sheet
  outputSheet.getRange(1, 1).setValue('d3.js Treemap JSON:');
  outputSheet.getRange(2, 1).setValue(jsonString);

  // Format the output
  outputSheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  outputSheet.getRange(2, 1).setWrap(true).setVerticalAlignment('top');
  outputSheet.setColumnWidth(1, 600);

  // Activate the output sheet
  ss.setActiveSheet(outputSheet);
}

/**
 * Escapes HTML special characters
 * @param {string} text - The text to escape
 * @return {string} The escaped text
 */
function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
