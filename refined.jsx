// Set Photoshop ruler units to inches
app.preferences.rulerUnits = Units.INCHES;

// Constants for tile dimensions, labels, and settings
const TILE_WIDTH_IN_INCHES = 6;
const TILE_HEIGHT_IN_INCHES = 4;
const DEFAULT_RESOLUTION = 300; // DPI for image processing
const LABEL_COLOR_RED = 255;
const LABEL_COLOR_GREEN = 0;
const LABEL_COLOR_BLUE = 0;
const PROGRESS_BAR_MESSAGE = 'Processing...';
const PROGRESS_BAR_CANCEL_LABEL = 'Cancel';
const LABEL_TEXT_SIZE = 6;
const LABEL_OPACITY = 40; // Percent
const LABEL_Y_OFFSET = 0.25; // Inches from top for label placement
const JPEG_QUALITY = 12;

// Get the active document (assume this is the original document already sized for the wall)
const originalDoc = app.activeDocument;

// Select folder to save tiles
const folder = Folder.selectDialog('Select folder to save the 6x4 inch tiles:');

// Get wall dimensions from the user
const wallWidth = parseFloat(prompt('Enter the width of the wall in inches: (Must be a multiple of 6 to fit 6x4 inch tiles without cutting)', ''));
const wallHeight = parseFloat(prompt('Enter the height of the wall in inches: (Must be a multiple of 4 to fit 6x4 inch tiles without cutting)', ''));

// Calculate the number of columns and rows needed to cover the wall
const cols = Math.ceil(wallWidth / TILE_WIDTH_IN_INCHES);
const rows = Math.ceil(wallHeight / TILE_HEIGHT_IN_INCHES);
const totalTiles = cols * rows;

// Create and display progress bar
const progressBar = createProgressBar(totalTiles);
var operationCancelled = false;

// Loop to process each tile
for (var c = 1; c <= cols; c++) {
  for (var r = 1; r <= rows; r++) {
    if (progressBar) progressBar.bar.value++;

    if (progressBar && progressBar.btnCancel.clicked) {
      originalDoc.close(SaveOptions.DONOTSAVECHANGES);
      operationCancelled = true;
      break;
    }

    // Calculate tile position
    var x = (c - 1) * TILE_WIDTH_IN_INCHES;
    var y = (r - 1) * TILE_HEIGHT_IN_INCHES;

    // Generate tile label using padded column and row numbers
    var colStr = padNumber(c, 2); // Pad column number to 2 digits
    var rowStr = padNumber(r, 2); // Pad row number to 2 digits
    var tileLabel = colStr + rowStr; // Resulting label will be like "0101", "0102", etc.
    var outputFile = new File(folder + '/' + tileLabel + '.jpg');

    // Duplicate original document for cropping each tile
    var tileDoc = originalDoc.duplicate();
    tileDoc.crop([x, y, x + TILE_WIDTH_IN_INCHES, y + TILE_HEIGHT_IN_INCHES]);

    // Add label to the tile
    var textLayer = tileDoc.artLayers.add();
    textLayer.kind = LayerKind.TEXT;
    textLayer.textItem.contents = tileLabel;
    var labelX = TILE_WIDTH_IN_INCHES / 2;
    var labelY = LABEL_Y_OFFSET; // Use original label Y offset
    textLayer.textItem.position = [labelX, labelY];
    textLayer.textItem.justification = Justification.CENTER;
    textLayer.textItem.color = getLabelColor(); // Use original color setting
    textLayer.textItem.size = LABEL_TEXT_SIZE; // Keep original text size
    textLayer.opacity = LABEL_OPACITY; // Keep original opacity

    // Save tile as JPEG
    var options = new JPEGSaveOptions();
    options.embedColorProfile = true;
    options.quality = JPEG_QUALITY; // Keep original JPEG quality
    tileDoc.saveAs(outputFile, options, true);
    tileDoc.close(SaveOptions.DONOTSAVECHANGES); // Close to free up memory
  }
  if (operationCancelled) {
    break;
  }
}

// Clean up and final messages
originalDoc.close(SaveOptions.DONOTSAVECHANGES);
progressBar.close();

// Alert for completion or cancellation
if (!operationCancelled) {
  alert('All tiles have been saved.');
} else {
  alert('Operation was cancelled. Some tiles might have been saved.');
}

// Helper functions

function createProgressBar(maxValue) {
  var w = new Window('palette', PROGRESS_BAR_MESSAGE, { x: 0, y: 0, width: 340, height: 60 });
  w.bar = w.add('progressbar', { x: 20, y: 12, width: 300, height: 12 }, 0, maxValue);
  w.btnCancel = w.add('button', { x: 130, y: 30, width: 80, height: 20 }, PROGRESS_BAR_CANCEL_LABEL);
  w.btnCancel.onClick = function () { w.close(); };
  w.show();
  return w;
}

function padNumber(num, size) {
  var s = num + '';
  while (s.length < size) s = '0' + s;
  return s;
}

function getLabelColor() {
  var color = new SolidColor();
  color.rgb.red = LABEL_COLOR_RED;
  color.rgb.green = LABEL_COLOR_GREEN;
  color.rgb.blue = LABEL_COLOR_BLUE;
  return color;
}
