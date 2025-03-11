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
    // document.getElementById("Convert-to-XML").onclick = convertPptxToXml;
    // document.getElementById("handle-file-upload").onclick = handleFileUpload;
    document.getElementById("uploadPptx").addEventListener("change", handleFileUpload);
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
            let slideFiles = Object.keys(zip.files).filter(file => file.startsWith("ppt/slides/") && file.endsWith(".xml"));

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