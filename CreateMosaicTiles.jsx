var originalDoc = app.activeDocument;
var folder = Folder.selectDialog("Select folder to save the 6x4 inch tiles:");

var widthInInches = 6;
var heightInInches = 4;

// Get wall dimensions or area from the user
var wallWidth = parseFloat(prompt("Enter the width of the wall (in inches):", ""));
var wallHeight = parseFloat(prompt("Enter the height of the wall (in inches):", ""));

// Calculate the scaling factor
var scaleFactor = Math.max(wallWidth / (originalDoc.width / originalDoc.resolution), wallHeight / (originalDoc.height / originalDoc.resolution));

// Duplicate the original document
var scaledDoc = originalDoc.duplicate();

// Scale the image while maintaining aspect ratio
scaledDoc.resizeImage(undefined, undefined, scaleFactor * originalDoc.resolution, ResampleMethod.BICUBIC);

// Resize the canvas to the wall dimensions
scaledDoc.resizeCanvas(wallWidth * originalDoc.resolution, wallHeight * originalDoc.resolution, AnchorPosition.MIDDLECENTER);

// Calculate the number of columns and rows based on the wall dimensions
var cols = Math.ceil(wallWidth / widthInInches);
var rows = Math.ceil(wallHeight / heightInInches);

var widthInPixels = widthInInches * originalDoc.resolution;
var heightInPixels = heightInInches * originalDoc.resolution;

for (var c = cols - 1; c >= 0; c--) {
  for (var r = rows - 1; r >= 0; r--) {
    var x = c * widthInPixels;
    var y = r * heightInPixels;

    var colStr = "col" + String("00" + (c + 1)).slice(-2);
    var rowStr = "row" + String("00" + (r + 1)).slice(-2);
    var outputFile = new File(folder + "/" + colStr + "_" + rowStr + ".jpg");

    // Duplicate the scaled document
    var newDoc = scaledDoc.duplicate();

    // Crop the duplicated document
    newDoc.crop([x, y, x + widthInPixels, y + heightInPixels]);

    // Save the cropped document as a JPEG
    var options = new JPEGSaveOptions();
    options.embedColorProfile = true;
    options.quality = 12; // Set the quality level (1-12)
    newDoc.saveAs(outputFile, options, true);

    // Close the cropped document without saving
    newDoc.close(SaveOptions.DONOTSAVECHANGES);
  }
}

// Close the scaled document without saving
scaledDoc.close(SaveOptions.DONOTSAVECHANGES);

alert("All tiles have been saved.");