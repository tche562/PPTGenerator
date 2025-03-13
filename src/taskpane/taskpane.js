/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("create-slide").onclick = function () {
      createTextFromXML(extractedTextData_test);
    };
    // document.getElementById("Convert-to-XML").onclick = convertPptxToXml;
    // document.getElementById("handle-file-upload").onclick = handleFileUpload;
    document.getElementById("upload-pptx").addEventListener("change", handleFileUpload);
  }
});

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

async function handleFileUpload(event) {
  if (!event.target.files.length) {
    console.error("No file selected.");
    return;
  }

  let file = event.target.files[0];

  // Validate PPTX file format
  if (!file.name.endsWith(".pptx")) {
    console.error("Invalid file format. Please upload a .pptx file.");
    return;
  }

  let reader = new FileReader();

  reader.onload = async function (e) {
    let arrayBuffer = e.target.result;

    try {
      let zip = await JSZip.loadAsync(arrayBuffer);
      console.log("Files in the PPTX:", Object.keys(zip.files));

      // Find all slide XML files dynamically
      let slideFiles = Object.keys(zip.files).filter(
        (file) => file.startsWith("ppt/slides/") && file.endsWith(".xml")
      );

      if (slideFiles.length > 0) {
        let slideXml = await zip.file(slideFiles[0]).async("text");
        console.log("First Slide XML Content:", slideXml);
      } else {
        console.error("No slides found in the PPTX.");
      }
    } catch (error) {
      console.error("Error processing PPTX:", error);
    }
  };

  reader.readAsArrayBuffer(file);
}

export async function createTextFromXML(data) {
  console.log("get in");

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;

    data.forEach((item) => {
      const textBox = shapes.addTextBox(item.text, {
        left: item.x,
        top: item.y,
        width: item.width,
        height: item.height,
      });
      textBox.name = `TextBox_${item.id}`;

      // Ensure text frame exists

      const textRange = textBox.textFrame.textRange;
      textRange.font.name = item.fontName || "Segoe UI"; // Default to Segoe UI if not specified
      textRange.font.size = item.fontSize || 16; // Default size 16pt
      textRange.font.bold = item.isBold || false;
      textRange.font.italic = item.isItalic || false;
      if (item.isBullet) {
        textRange.paragraphFormat.bulletFormat.visible = true; // Enable bullets
      }
    });

    return context.sync();
  });
}

const extractedTextData = [
  {
    id: 1,
    text: "Creating a mind map",
    x: 444500 / 9525,
    y: 412137 / 9525,
    width: 9146972 / 9525,
    height: 640080 / 9525,
    fontName: "Segoe UI Semibold",
    fontSize: 28,
    isBold: true,
    isBullet: false, // Not a bullet point
  },
  {
    id: 2,
    text: "Mind maps are a great way to:\n •  Channel your creativity\n •  Generate ideas\n •  See visual relationships",
    x: 444500 / 9525,
    y: 1509612 / 9525,
    width: 4975869 / 9525,
    height: 3977640 / 9525,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: false,
  },
];

export const extractedTextData_test = [
  {
    id: 1,
    text: "Creating a mind map",
    x: 444500 / 9525,
    y: 412137 / 9525,
    width: 9146972 / 9525,
    height: 640080 / 9525,
    fontName: "Segoe UI Semibold",
    fontSize: 28,
    isBold: true,
    isBullet: false, // Not a bullet point
  },
  {
    id: 2,
    text: "Mind maps are a great way to:",
    x: 444500 / 9525,
    y: 1509612 / 9525,
    width: 4975869 / 9525,
    height: 3977640 / 9525,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: false,
  },
  {
    id: 3,
    text: "  Channel your creativity",
    x: 464500 / 9525,
    y: (1509612 + 300000) / 9525,
    width: 4975869 / 9525,
    height: 3977640 / 9525,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
  },
  {
    id: 4,
    text: "  Generate ideas",
    x: 464500 / 9525,
    y: (1509612 + 600000) / 9525,
    width: 4975869 / 9525,
    height: 3977640 / 9525,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
  },
  {
    id: 5,
    text: "  See visual relationships",
    x: 464500 / 9525,
    y: (1509612 + 900000) / 9525,
    width: 4975869 / 9525,
    height: 3977640 / 9525,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
  },
  {
    id: 6,
    text: "  Improve your memory",
    x: 464500 / 9525,
    y: (1509612 + 1200000) / 9525,
    width: 4975869 / 9525,
    height: 3977640 / 9525,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
  },
];
