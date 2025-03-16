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
    document.getElementById("create-text-box-1").onclick = function () {
      createTextFromXML(extractedTextData_test);
    };
    document.getElementById("create-text-box-2").onclick = function () {
      createTextFromXML(extractedTextData_test1);
    };
    document.getElementById("create-round").onclick = function () {
      createShapeFromXML(roundData1);
    };
    document.getElementById("create-line").onclick = function () {
      createLineFromXML(lineData);
    };
    document.getElementById("create-image").onclick = function () {
      createImage(svgData);
    };
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

      // Define specific files to extract and log
      const targetFiles = [
        "ppt/theme/theme21.xml",
        "ppt/slides/slide21.xml",
        "ppt/slides/_rels/slide21.xml.rels",
        "ppt/slideLayouts/slideLayout21.xml",
        "ppt/slideLayouts/_rels/slideLayout21.xml.rels",
      ];

      for (let file of targetFiles) {
        if (zip.files[file]) {
          let fileContent = await zip.file(file).async("text");
          console.log(`Content of ${file}:`, fileContent);
        } else {
          console.warn(`File not found: ${file}`);
        }
      }
      const svgFiles = [
        "ppt/media/image3.svg",
        "ppt/media/image53.svg",
        "ppt/media/image72.svg",
        "ppt/media/image94.svg",
      ];

      for (let file of svgFiles) {
        if (zip.files[file]) {
          const svgData = await zip.file(file).async("string");
          console.log(`SVG extracted successfully! Content of ${file}:`, svgData);
        } else {
          console.warn(`SVG file not found: ${file}`);
        }
      }
    } catch (error) {
      console.error("Error processing PPTX:", error);
    }
  };

  reader.readAsArrayBuffer(file);
}

export async function createTextFromXML(data) {
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
      textRange.font.color = item.fontColor || "black";
      if (item.isBullet) {
        textRange.paragraphFormat.bulletFormat.visible = true; // Enable bullets
      }
      if (item.horizontalAlignment != null || undefined) {
        textRange.paragraphFormat.horizontalAlignment = item.horizontalAlignment;
      }
    });

    return context.sync();
  });
}

export async function createShapeFromXML(data) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;

    data.forEach((item) => {
      const shapeOptions = {
        left: item.x,
        top: item.y,
        height: item.height,
        width: item.width,
      };
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, shapeOptions);
      shape.name = `Shape_${item.id}`;

      // Apply the properties to the shape
      shape.rotation = item.rotation;

      console.log(item.fillColor);

      if (item.fillColor !== "N/A") {
        shape.fill.setSolidColor(item.fillColor);
      }
      shape.lineFormat.visible = false;
    });

    return context.sync();
  });
}

export async function createLineFromXML(data) {
  // This function gets the collection of shapes on the first slide,
  // and adds a line to the collection, while specifying its
  // start and end points. Then it names the shape.
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;

    // For a line, left and top are the coordinates of the start point,
    // while height and width are the coordinates of the end point.
    data.forEach((item) => {
      const shapeOptions = {
        left: item.x,
        top: item.y,
        height: item.height,
        width: item.width,
      };
      const line = shapes.addLine(PowerPoint.ConnectorType.straight, shapeOptions);
      line.name = `Line_${item.id}`;
      line.lineFormat.color = item.color;
      line.lineFormat.weight = 2;
    });
    return context.sync();
  });
}

export function createImage(data) {
  data.forEach((item) => {
    console.log(item.svg);
    
    Office.context.document.setSelectedDataAsync(
      item.svg,
      {
        coercionType: Office.CoercionType.XmlSvg,
        imageLeft: item.x,
        imageTop: item.y,
        imageWidth: item.width,
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      }
    );
  });
}

const roundData1 = [
  {
    shapeId: 5,
    type: "Ellipse",
    x: 6191627 / 12700,
    y: 2349747 / 12700,
    width: 1000125 / 12700,
    height: 1000125 / 12700,
    rotation: 0,
    fillColor: "#4472C4",
    text: "Oval 4",
  },
  {
    shapeId: 15,
    type: "Ellipse",
    x: 9681281 / 12700,
    y: 4575965 / 12700,
    width: 1000125 / 12700,
    height: 1000125 / 12700,
    rotation: 0,
    fillColor: "#9B5AC8",
    text: "Oval 14",
  },
  {
    shapeId: 24,
    type: "Ellipse",
    x: 10362310 / 12700,
    y: 2349747 / 12700,
    width: 1000125 / 12700,
    height: 1000125 / 12700,
    rotation: 0,
    fillColor: "#D24726",
    text: "Oval 23",
  },
  {
    shapeId: 69,
    type: "Ellipse",
    x: 7960738 / 12700,
    y: 1939633 / 12700,
    width: 1630734 / 12700,
    height: 1630734 / 12700,
    rotation: 0,
    fillColor: "#3B3838",
    text: "Oval 68",
  },
  {
    shapeId: 9,
    type: "Ellipse",
    x: 7081628 / 12700,
    y: 4258095 / 12700,
    width: 1000125 / 12700,
    height: 1000125 / 12700,
    rotation: 0,
    fillColor: "#70AD47",
    text: "Oval 8",
  },
];

const lineData = [
  {
    shapeId: 38,
    type: "Line",
    x: 9181218 / 12700,
    y: 3429000 / 12700,
    width: 836715 / 12700,
    height: 1277695 / 12700,
    color: "#000000",
  },
  {
    shapeId: 43,
    type: "Line",
    x: 9427475 / 12700,
    y: 2769592 / 12700,
    width: 934835 / 12700,
    height: 80218 / 12700,
    color: "#000000",
  },
  {
    shapeId: 30,
    type: "Line",
    x: (7188401 + 768986) / 12700,
    y: 2785215 / 12700,
    width: -768986 / 12700,
    height: 88013 / 12700,
    color: "#000000",
  },
  {
    shapeId: 31,
    type: "Line",
    x: (7742360 + 552624) / 12700,
    y: 3349872 / 12700,
    width: -552624 / 12700,
    height: 945903 / 12700,
    color: "#000000",
  },
  {
    shapeId: 12,
    type: "Line",
    x: 533400 / 12700,
    y: 1104900 / 12700,
    width: 11119104 / 12700,
    height: 0.00000001,
    color: "#D24726",
  },
];

const extractedTextData = [
  {
    id: 1,
    text: "Creating a mind map",
    x: 444500 / 12700,
    y: 412137 / 12700,
    width: 9146972 / 12700,
    height: 640080 / 12700,
    fontName: "Segoe UI Semibold",
    fontSize: 28,
    isBold: true,
    isBullet: false, // Not a bullet point
  },
  {
    id: 2,
    text: "Mind maps are a great way to:\n •  Channel your creativity\n •  Generate ideas\n •  See visual relationships",
    x: 444500 / 12700,
    y: 1509612 / 12700,
    width: 4975869 / 12700,
    height: 3977640 / 12700,
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
    x: 444500 / 12700,
    y: 412137 / 12700,
    width: 9146972 / 12700,
    height: 640080 / 12700,
    fontName: "Segoe UI Semibold",
    fontSize: 28,
    isBold: true,
    isBullet: false,
    horizontalAlignment: "",
  },
  {
    id: 2,
    text: "Mind maps are a great way to:",
    x: 444500 / 12700,
    y: 1509612 / 12700,
    width: 4975869 / 12700,
    height: 3977640 / 12700,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: false,
    horizontalAlignment: "",
  },
  {
    id: 3,
    text: "  Channel your creativity",
    x: 464500 / 12700,
    y: (1509612 + 300000) / 12700,
    width: 4975869 / 12700,
    height: 3977640 / 12700,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
    horizontalAlignment: "",
  },
  {
    id: 4,
    text: "  Generate ideas",
    x: 464500 / 12700,
    y: (1509612 + 600000) / 12700,
    width: 4975869 / 12700,
    height: 3977640 / 12700,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
    horizontalAlignment: "",
  },
  {
    id: 5,
    text: "  See visual relationships",
    x: 464500 / 12700,
    y: (1509612 + 900000) / 12700,
    width: 4975869 / 12700,
    height: 3977640 / 12700,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
    horizontalAlignment: "",
  },
  {
    id: 6,
    text: "  Improve your memory",
    x: 464500 / 12700,
    y: (1509612 + 1200000) / 12700,
    width: 4975869 / 12700,
    height: 3977640 / 12700,
    fontName: "Segoe UI",
    fontSize: 16,
    isBold: false,
    isBullet: true,
    horizontalAlignment: "",
  },
];

export const extractedTextData_test1 = [
  {
    id: 50,
    text: "Topic",
    x: 6083842 / 12700,
    y: 3410881 / 12700,
    width: 1198880 / 12700,
    height: 369332 / 12700,
    fontName: "Segoe UI(Body)",
    fontSize: 18,
    fontColor: "#000000", // Default black
    isBold: false,
    isBullet: false,
    horizontalAlignment: "center",
  },
  {
    id: 51,
    text: "Audience",
    x: 6973454 / 12700,
    y: 5320732 / 12700,
    width: 1198880 / 12700,
    height: 369332 / 12700,
    fontName: "Segoe UI(Body)",
    fontSize: 18,
    fontColor: "#000000",
    isBold: false,
    isBullet: false,
    horizontalAlignment: "center",
  },
  {
    id: 52,
    text: "Visuals",
    x: 9591472 / 12700,
    y: 5638602 / 12700,
    width: 1198880 / 12700,
    height: 369332 / 12700,
    fontName: "Segoe UI(Body)",
    fontSize: 18,
    fontColor: "#000000",
    isBold: false,
    isBullet: false,
    horizontalAlignment: "center",
  },
  {
    id: 53,
    text: "Schedule",
    x: 10017933 / 12700,
    y: 3412384 / 12700,
    width: 1608324 / 12700,
    height: 369332 / 12700,
    fontName: "Segoe UI(Body)",
    fontSize: 18,
    fontColor: "#000000",
    isBold: false,
    isBullet: false,
    horizontalAlignment: "center",
  },
  {
    id: 3,
    text: "Product pitch",
    x: 8100607 / 12700,
    y: 2446426 / 12700,
    width: 1326868 / 12700,
    height: 646331 / 12700,
    fontName: "Segoe UI(Body)",
    fontSize: 18,
    fontColor: "#FFFFFF",
    isBold: false,
    isBullet: false,
    horizontalAlignment: "center",
  },
];

const svgData = [
  {
    x: 9824408 / 12700,
    y: 4723235 / 12700,
    width: 734624 / 12700,
    svg: '<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" id="Icons_Palette" overflow="hidden"><style>.MsftOfcThm_Background1_Fill_v2 {fill:#FFFFFF; }</style><path d="M78 51C75.2 51 73 49.4 73 47.5 73 45.6 75.2 44 78 44 80.8 44 83 45.6 83 47.5 83 49.4 80.8 51 78 51ZM73 62C70.2 62 68 60.4 68 58.5 68 56.6 70.2 55 73 55 75.8 55 78 56.6 78 58.5 78 60.4 75.8 62 73 62ZM62 34C59.2 34 57 32.4 57 30.5 57 28.6 59.2 27 62 27 64.8 27 67 28.6 67 30.5 67 32.4 64.8 34 62 34ZM62 68C59.2 68 57 66.4 57 64.5 57 62.6 59.2 61 62 61 64.8 61 67 62.6 67 64.5 67 66.4 64.8 68 62 68ZM48 71C45.2 71 43 69.4 43 67.5 43 65.6 45.2 64 48 64 50.8 64 53 65.6 53 67.5 53 69.4 50.8 71 48 71ZM42.6 30.6C44.2 29 46.3 28.7 47.4 29.8 48.5 30.9 48.1 33 46.6 34.6 45 36.2 42.9 36.5 41.8 35.4 40.6 34.3 41 32.1 42.6 30.6ZM34 69C31.2 69 29 67.4 29 65.5 29 63.6 31.2 62 34 62 36.8 62 39 63.6 39 65.5 39 67.4 36.8 69 34 69ZM73 34C75.8 34 78 35.6 78 37.5 78 39.4 75.8 41 73 41 70.2 41 68 39.4 68 37.5 68 35.6 70.2 34 73 34ZM48 20C25.3 20 28.9 28.9 31 31L36 36C39.2 39.3 35.3 43 31.6 41.4L19.1 36C10.3 32.2 8 42.9 8 48 8 63.5 25.9 76 48 76 70.1 76 88 63.5 88 48 88 32.5 70.1 20 48 20Z" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" fill="#FFFFFF" fill-opacity="1"/></svg>',
  },
  {
    x: 6352151 / 12700,
    y: 2505048 / 12700,
    width: 679076 / 12700,
    svg: '<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" id="Icons_Network" overflow="hidden"><style>.MsftOfcThm_Background1_Fill_v2 {fill:#FFFFFF; }</style><path d="M85 37.4C83.3 33.3 78.6 31.4 74.5 33.1 71.1 34.5 69.1 38.1 69.6 41.6L56.7 47C55.3 44.6 52.8 42.8 50 42.3L50 28.4C53.4 27.5 56 24.4 56 20.7 56 16.3 52.4 12.7 48 12.7L48 12.7C43.6 12.7 40 16.3 40 20.7 40 24.4 42.6 27.5 46 28.4L46 42.2C43.1 42.7 40.7 44.5 39.3 46.9L26.4 41.5C26.9 38 25 34.4 21.5 33 17.4 31.3 12.7 33.2 11 37.3 9.3 41.4 11.2 46.1 15.3 47.8 18.7 49.2 22.6 48.1 24.7 45.2L37.9 50.6C37.8 51 37.8 51.5 37.8 51.9 37.8 54.1 38.5 56.2 39.8 57.9L29.4 68.4C26.3 66.6 22.3 67 19.7 69.6 16.6 72.7 16.6 77.8 19.7 80.9 22.8 84 27.9 84 31 80.9 33.6 78.3 34 74.3 32.2 71.2L42.8 60.6C44.3 61.5 46 62 47.8 62 47.9 62 47.9 62 48 62 48.1 62 48.1 62 48.2 62 50 62 51.7 61.5 53.2 60.6L63.8 71.2C62 74.3 62.4 78.3 65 81 68.1 84.1 73.2 84.1 76.3 81 79.4 77.9 79.4 72.8 76.3 69.7 73.7 67.1 69.7 66.7 66.6 68.5L56.2 58C57.5 56.3 58.2 54.3 58.2 52 58.2 51.6 58.2 51.1 58.1 50.7L71.3 45.3C73.4 48.1 77.3 49.3 80.7 47.9 84.7 46.1 86.7 41.5 85 37.4Z" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" fill="#FFFFFF" fill-opacity="1"/></svg>',
  },
  {
    x: 10502179 / 12700,
    y: 2506513 / 12700,
    width: 724630 / 12700,
    svg: '<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" id="Icons_Team" overflow="hidden"><style>.MsftOfcThm_Background1_Fill_v2 {fill:#FFFFFF;}</style><g><circle cx="73" cy="16.2" r="7" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" fill="#FFFFFF" fill-opacity="1"/><circle cx="23" cy="16.2" r="7" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" fill="#FFFFFF" fill-opacity="1"/><circle cx="48" cy="16.2" r="7" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" fill="#FFFFFF" fill-opacity="1"/><path d="M65.8 48.8 61.3 32.2C61.1 31.6 60.8 31 60.4 30.6 58.5 28.6 56.1 27.1 53.5 26.2 51.8 25.6 50 25.3 48.1 25.3 46.2 25.3 44.4 25.6 42.7 26.2 40 27.1 37.7 28.6 35.8 30.6 35.4 31.1 35.1 31.6 34.9 32.2L30.4 48.8C30 50.4 30.8 52.2 32.5 52.6 32.8 52.7 33 52.7 33.3 52.7 34.6 52.7 35.8 51.8 36.2 50.5L40.2 35.9 40.2 86.8 46.2 86.8 46.2 58.2 50.2 58.2 50.2 86.7 56.2 86.7 56.2 35.9 60.2 50.5C60.6 51.8 61.8 52.7 63.1 52.7 63.4 52.7 63.6 52.7 63.9 52.6 65.3 52.2 66.2 50.4 65.8 48.8Z" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" fill="#FFFFFF" fill-opacity="1"/><path d="M28.3 48.3 32.8 31.7C33 30.8 33.5 30 34 29.3 32.4 27.9 30.4 26.8 28.3 26.1 26.6 25.5 24.8 25.2 22.9 25.2 21 25.2 19.2 25.5 17.5 26.1 14.8 27 12.5 28.5 10.6 30.5 10.2 31 9.9 31.5 9.7 32.1L5.2 48.8C4.8 50.4 5.6 52.2 7.3 52.6 7.6 52.7 7.8 52.7 8.1 52.7 9.4 52.7 10.6 51.8 11 50.5L15 35.9 15 86.8 21 86.8 21 58.2 25 58.2 25 86.7 31 86.7 31 54.3C28.9 53.3 27.7 50.8 28.3 48.3Z" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" fill="#FFFFFF" fill-opacity="1"/><path d="M90.8 48.8 86.2 32.2C86 31.6 85.7 31 85.3 30.6 83.4 28.6 81 27.1 78.4 26.2 76.7 25.6 74.9 25.3 73 25.3 71.1 25.3 69.3 25.6 67.6 26.2 65.5 26.9 63.6 28 61.9 29.4 62.5 30.1 62.9 30.9 63.1 31.7L67.6 48.3C68.3 50.8 67 53.3 64.9 54.3L64.9 86.8 70.9 86.8 70.9 58.2 74.9 58.2 74.9 86.7 80.9 86.7 80.9 35.9 84.9 50.5C85.3 51.8 86.5 52.7 87.8 52.7 88.1 52.7 88.3 52.7 88.6 52.6 90.3 52.2 91.2 50.4 90.8 48.8Z" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" fill="#FFFFFF" fill-opacity="1"/></g></svg>',
  },
  {
    x: 7227805 / 12700,
    y: 4412120 / 12700,
    width: 692074 / 12700,
    svg: '<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" id="Icons_Puzzle" overflow="hidden"><style>.MsftOfcThm_Background1_Fill_v2 { fill:#FFFFFF; }</style><path d="M59.7 66.7C53.1 66.9 50.7 58.4 55.4 53.5L56.1 52.8C61 48.1 69.7 50.3 69.5 56.9 69.4 60.7 73.9 65.3 76.6 62.6L88 51.2 71 34.2C68.3 31.5 72.9 27 76.7 27.1 83.3 27.3 85.5 18.6 80.8 13.7L80.1 13C75.2 8.3 66.7 10.7 66.9 17.3 67 21.1 62.5 25.7 59.8 23L42.8 6 31.3 17.4C28.6 20.1 33.2 24.6 37 24.5 43.6 24.3 46 32.8 41.3 37.7L40.6 38.4C35.7 43.1 27 40.9 27.2 34.3 27.3 30.5 22.8 25.9 20.1 28.6L8 40.8 25 57.8C27.7 60.5 23.1 65 19.3 64.9 12.7 64.7 10.5 73.4 15.2 78.3L15.9 79C20.8 83.7 29.3 81.3 29.1 74.7 29 70.9 33.5 66.3 36.2 69L53.2 86 65.4 73.8C68.1 71.1 63.6 66.6 59.7 66.7Z" class="MsftOfcThm_Background1_Fill_v2" stroke="none" stroke-width="1" stroke-linecap="butt" fill="#FFFFFF" fill-opacity="1"/></svg>',
  },
];

