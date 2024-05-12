/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}


// Function to convert RGB color to hexadecimal format
function rgbToHex(rgb) {
  // Extract the RGB values from the string
  const [r, g, b] = rgb.match(/\d+/g);

  // Convert each RGB value to hexadecimal
  const hexR = Number(r)
    .toString(16)
    .padStart(2, "0");
  const hexG = Number(g)
    .toString(16)
    .padStart(2, "0");
  const hexB = Number(b)
    .toString(16)
    .padStart(2, "0");

  // Return the hexadecimal color string
  return `#${hexR}${hexG}${hexB}`;
}

Office.onReady(function () {
  // Office is ready.
  $(document).ready(function () {
    // The document is ready.
    $(".square").on("click", function() {
  const type = $(this).data("type");
  switch (type) {
    case "fill":
      tryCatch(() => applyFillColor(event));
      break;
    case "outline":
      tryCatch(() => applyOutlineColor(event));
      break;
    default:
      break;
  }
});
$(".fa-font").on("click", (event) => tryCatch(() => applyTextColor(event))); // Get Selected Text and apply Color
  });
});

// From .square get background color
// From .fa-font get color

// Function to apply color to a shape in PowerPoint using the Office JavaScript API
async function applyFillColor(event) {
  await PowerPoint.run(async (context) => {
    const clickedElement = event.target;
    const clickedElementId = event.target.id;
    // const shapes = context.presentation.getSelectedShapes();
    console.log("Element ID: ", clickedElementId);
    let hexColor;
    // If it does not contain the id remove
    if (clickedElementId === "remove-fill") {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items"); // Load the 'items' property
      await context.sync();

      shapes.items.forEach((shape) => {
        shape.fill.clear();
      });
    }
    // Else Execute normally

    // Get background color of clicked element
    else {
      const squareStyles = window.getComputedStyle(clickedElement);
      const bgColor = squareStyles.getPropertyValue("background-color");
      // Convert background color to hexadecimal
      hexColor = rgbToHex(bgColor);
    }
    // Get the selected shapes in the presentation
    const shapes = context.presentation.getSelectedShapes();
    console.log("Selected Color: ", hexColor);
    // Get the shape you want to apply the color to (replace 'shapeName' with the actual name of the shape)
    await context.sync();

    // Apply the color to the shape

    //console.log("The selected shape is: ",shape);
    if (hexColor != null) {
      shapes.load("items");
      await context.sync();
      shapes.items.map((shape) => {
        shape.fill.setSolidColor(hexColor);
      });
    }
    shapes.items.forEach((shape) => {
      console.log("Shape ID: ", shape.id);
    });

    await context.sync();
  });
}

async function applyOutlineColor(event) {
  await PowerPoint.run(async (context) => {
    // Get background color of clicked element
    const clickedElement = event.target;
    console.log("Element ID: ", clickedElement.id);
    let hexColor;
    if ($(clickedElement).hasClass("white")) {
      hexColor = "white";
    } else if (clickedElement.id === "remove-outline") {
      hexColor = null;
    } else {
      const innerSquareStyles = window.getComputedStyle(clickedElement);
      const borderColor = innerSquareStyles.getPropertyValue("border-color");

      // Convert background color to hexadecimal
      hexColor = rgbToHex(borderColor);
    }
    // Get the selected shapes in the presentation
    const shapes = context.presentation.getSelectedShapes();
    console.log("Selected Color: ", hexColor);
    // Get the shape you want to apply the color to (replace 'shapeName' with the actual name of the shape)
    await context.sync();

    // Apply the color to the shape
    // shapes.fill.setSolidColor(color);
    shapes.load("items");
    await context.sync();
    //console.log("The selected shape is: ",shape);
    shapes.items.map((shape) => {
      shape.lineFormat.color = hexColor;
    });
    shapes.items.forEach((shape) => {
      console.log("Shape ID: ", shape.id);
    });
    // Get the shape you want to apply the color to (replace 'shapeName' with the actual name of the shape)
    await context.sync();
  });
}

async function applyTextColor(event) {
  await PowerPoint.run(async (context) => {
    const clickedElement = event.target;
    console.log("Element ID: ", clickedElement.id);
    let hexColor;
    if (clickedElement.id === "remove-text") {
      hexColor = null;
    } else {
      const squareStyles = window.getComputedStyle(clickedElement);
      const bgColor = squareStyles.getPropertyValue("color");
      // Convert background color to hexadecimal
      hexColor = rgbToHex(bgColor);
    }
    // Get the selected shapes in the presentation
    const shapes = context.presentation.getSelectedShapes();
    console.log("Selected Color: ", hexColor);
    // Get the shape you want to apply the color to (replace 'shapeName' with the actual name of the shape)
    await context.sync();

    // Apply the color to the shape
    // shapes.fill.setSolidColor(color);
    shapes.load("items");
    await context.sync();
    //console.log("The selected shape is: ",shape);
    shapes.items.map((shape) => {
      shape.textFrame.textRange.font.color = hexColor;
    });
    shapes.items.forEach((shape) => {
      console.log("Shape ID: ", shape.id);
    });
    // Get the shape you want to apply the color to (replace 'shapeName' with the actual name of the shape)
    await context.sync();
  });
}

