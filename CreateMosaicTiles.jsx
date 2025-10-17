// Constants for easy configuration
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

// Function to convert column index to letter
function columnToLetter(column) {
  var letterIndex,
    letter = '';
  while (column > 0) {
    letterIndex = (column - 1) % 26;
    letter = String.fromCharCode(letterIndex + 65) + letter;
    column = (column - letterIndex - 1) / 26;
  }
  return letter;
}

// Function to get label color
function getLabelColor() {
  const labelColor = new SolidColor();
  labelColor.rgb.red = LABEL_COLOR_RED;
  labelColor.rgb.green = LABEL_COLOR_GREEN;
  labelColor.rgb.blue = LABEL_COLOR_BLUE;
  return labelColor;
}

// Function to create a progress bar
function createProgressBar(totalTiles) {
  const win = new Window('palette', 'Processing Tiles', undefined, {
    closeButton: false,
  });
  win.bar = win.add(
    'progressbar',
    { x: 20, y: 12, width: 300, height: 20 },
    0,
    totalTiles,
  );
  win.stMessage = win.add(
    'statictext',
    { x: 20, y: 36, width: 300, height: 20 },
    PROGRESS_BAR_MESSAGE,
  );
  win.btnCancel = win.add(
    'button',
    { x: 125, y: 60, width: 90, height: 20 },
    PROGRESS_BAR_CANCEL_LABEL,
  );
  win.show();
  return win;
}

// Function to pad number with leading zeros
function padNumber(num, size) {
  const s = '00' + num;
  return s.substr(s.length - size);
}

// Set Photoshop ruler units to inches
app.preferences.rulerUnits = Units.INCHES;

// Resize original document to specified resolution
const originalDoc = app.activeDocument;
originalDoc.resizeImage(
  undefined,
  undefined,
  DEFAULT_RESOLUTION,
  ResampleMethod.NONE,
);

// Folder selection for saving tiles
const folder = Folder.selectDialog(
  'Select folder to save the 6x4 inch tiles:',
);

// Prompt user for wall dimensions
const wallWidth = parseFloat(
  prompt(
    'Enter the width of the wall in inches: (Must be a multiple of 6 to fit 6x4 inch tiles without cutting)',
    '',
  ),
);
const wallHeight = parseFloat(
  prompt(
    'Enter the height of the wall in inches: (Must be a multiple of 4 to fit 6x4 inch tiles without cutting)',
    '',
  ),
);

// Calculate scale factor for image
const scaleFactor = Math.min(
  wallWidth / originalDoc.width,
  wallHeight / originalDoc.height,
);

// Duplicate and resize original document
const scaledDoc = originalDoc.duplicate();
scaledDoc.resizeImage(
  undefined,
  undefined,
  scaleFactor * originalDoc.resolution,
  ResampleMethod.BICUBIC,
);

// Create new document for wall with white background
const wallDoc = app.documents.add(
  wallWidth,
  wallHeight,
  originalDoc.resolution,
  'Wall',
  NewDocumentMode.RGB,
  DocumentFill.WHITE,
);

// Copy and paste scaled image into wall document
app.activeDocument = scaledDoc;
scaledDoc.selection.selectAll();
scaledDoc.selection.copy();
app.activeDocument = wallDoc;
wallDoc.paste();
scaledDoc.close(SaveOptions.DONOTSAVECHANGES);

// Calculate number of columns and rows
const cols = Math.ceil(wallWidth / TILE_WIDTH_IN_INCHES);
const rows = Math.ceil(wallHeight / TILE_HEIGHT_IN_INCHES);
const totalTiles = cols * rows;

// Create and display progress bar
const progressBar = createProgressBar(totalTiles);
var operationCancelled = false;

// Loop to process each tile
for (var c = cols - 1; c >= 0; c--) {
  for (var r = rows - 1; r >= 0; r--) {
    if (progressBar) progressBar.bar.value++;

    if (progressBar && progressBar.btnCancel.clicked) {
      wallDoc.close(SaveOptions.DONOTSAVECHANGES);
      operationCancelled = true;
      break;
    }

    // Calculate tile position
    var x = c * TILE_WIDTH_IN_INCHES;
    var y = r * TILE_HEIGHT_IN_INCHES;

    // Generate tile label
    var colStr = columnToLetter(c + 1);
    var rowStr = padNumber(r + 1, 2); // Custom padding function
    var tileLabel = colStr + rowStr;
    var outputFile = new File(folder + '/' + tileLabel + '.jpg');

    // Duplicate wall document for cropping
    var tileDoc = wallDoc.duplicate();
    tileDoc.crop([x, y, x + TILE_WIDTH_IN_INCHES, y + TILE_HEIGHT_IN_INCHES]);

    // Add label to the tile
    var textLayer = tileDoc.artLayers.add();
    textLayer.kind = LayerKind.TEXT;
    textLayer.textItem.contents = tileLabel;
    var labelX = TILE_WIDTH_IN_INCHES / 2;
    var labelY = LABEL_Y_OFFSET; // Label position
    textLayer.textItem.position = [labelX, labelY];
    textLayer.textItem.justification = Justification.CENTER;
    textLayer.textItem.color = getLabelColor();
    textLayer.textItem.size = LABEL_TEXT_SIZE;
    textLayer.opacity = LABEL_OPACITY;

    // Save tile as JPEG
    var options = new JPEGSaveOptions();
    options.embedColorProfile = true;
    options.quality = JPEG_QUALITY;
    tileDoc.saveAs(outputFile, options, true);
    tileDoc.close(SaveOptions.DONOTSAVECHANGES);
  }
  if (operationCancelled) {
    break;
  }
}

// Clean up and final messages
wallDoc.close(SaveOptions.DONOTSAVECHANGES);
progressBar.close();

// Alert for completion or cancellation
if (!operationCancelled) {
  alert('All tiles have been saved.');
} else {
  alert('Operation was cancelled. Some tiles might have been saved.');
}
