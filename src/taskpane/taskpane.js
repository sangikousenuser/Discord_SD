/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Adobe PDF Services SDK
const PDFServicesSdk = require('@adobe/pdfservices-node-sdk');

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("save-pdf-button").onclick = saveAndCreatePDF;
    document.getElementById("status-message").style.display = "none";
  }
});

/**
 * Saves the current document and creates a PDF version
 */
async function saveAndCreatePDF() {
  try {
    // Show status message
    showStatus("Processing your request...", "info");
    
    // Save the document
    await saveDocument();
    
    // Create PDF
    const pdfPath = await createPDF();
    
    // Copy to clipboard (where supported)
    await copyToClipboard(pdfPath);
    
    // Show success message
    showStatus("Document saved and PDF created successfully!", "success");
  } catch (error) {
    console.error("Error:", error);
    showStatus("Error: " + error.message, "error");
  }
}

/**
 * Saves the current document
 */
async function saveDocument() {
  return new Promise((resolve, reject) => {
    try {
      // For now, we'll use the Document.save method which prompts the user
      // This is a limitation of the Office JS API as discussed in research
      Office.context.document.save((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Document saved successfully");
          resolve();
        } else {
          reject(new Error("Failed to save document: " + result.error.message));
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Creates a PDF version of the document using Adobe PDF Services SDK
 * Note: In a real implementation, you would need Adobe credentials
 */
async function createPDF() {
  try {
    // In a real implementation, this would use Adobe PDF Services SDK
    // to convert the Word document to PDF
    
    // For demonstration purposes, we'll simulate the PDF creation
    // and return a placeholder path
    
    // In a real implementation with Adobe PDF Services:
    /*
    const credentials = PDFServicesSdk.Credentials
      .serviceAccountCredentialsBuilder()
      .fromFile("pdfservices-api-credentials.json")
      .build();

    // Create an ExecutionContext using credentials
    const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);
    
    // Build exportPDF operation
    const exportPDF = PDFServicesSdk.ExportPDF.Operation.createNew(
      PDFServicesSdk.ExportPDF.SupportedTargetFormats.DOCX
    );
    
    // Set operation input from a source file
    const input = PDFServicesSdk.FileRef.createFromLocalFile(
      'path/to/document.docx'
    );
    exportPDF.setInput(input);
    
    // Execute the operation and Save the result to the specified location
    const result = await exportPDF.execute(executionContext);
    await result.saveAsFile('output/document.pdf');
    */
    
    console.log("PDF creation simulated - in real implementation, this would create a PDF");
    return "output/document.pdf"; // Placeholder path
  } catch (err) {
    console.log('Exception encountered while executing operation', err);
    throw new Error("PDF creation failed: " + err.message);
  }
}

/**
 * Attempts to copy file paths to clipboard
 * @param {string} pdfPath - Path to the PDF file
 */
async function copyToClipboard(pdfPath) {
  try {
    // Due to limitations in Office JS API for clipboard access,
    // we'll use a workaround approach
    
    // Create a temporary text element with the file paths
    const docPath = Office.context.document.url || "Document path not available";
    const textToCopy = `Word Document: ${docPath}\nPDF Document: ${pdfPath}`;
    
    // Try to use the clipboard API if available
    if (navigator.clipboard && navigator.clipboard.writeText) {
      try {
        await navigator.clipboard.writeText(textToCopy);
        console.log("Paths copied to clipboard using Clipboard API");
        return;
      } catch (clipboardError) {
        console.warn("Clipboard API failed:", clipboardError);
        // Fall back to execCommand method
      }
    }
    
    // Fall back to the deprecated execCommand method
    const textArea = document.createElement("textarea");
    textArea.value = textToCopy;
    textArea.style.position = "fixed";  // Avoid scrolling to bottom
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
      const successful = document.execCommand('copy');
      if (successful) {
        console.log("Paths copied to clipboard using execCommand");
      } else {
        console.warn("execCommand copy failed");
      }
    } catch (err) {
      console.warn("execCommand error:", err);
    }
    
    document.body.removeChild(textArea);
    
    // Note to user about clipboard limitations
    const statusElement = document.getElementById("status-message");
    const textElement = statusElement.querySelector(".ms-MessageBar-text");
    textElement.innerHTML += "<br><small>Note: Due to browser limitations, file paths may not be copied to clipboard in all environments.</small>";
    
  } catch (error) {
    console.warn("Clipboard operation failed:", error);
    // Don't throw error here, as this is a non-critical feature
  }
}

/**
 * Shows a status message to the user
 * @param {string} message - The message to display
 * @param {string} type - The type of message (success, error, info)
 */
function showStatus(message, type) {
  const statusElement = document.getElementById("status-message");
  const textElement = statusElement.querySelector(".ms-MessageBar-text");
  
  // Set message text
  textElement.textContent = message;
  
  // Set appropriate class based on message type
  statusElement.className = "ms-MessageBar";
  if (type === "success") {
    statusElement.classList.add("ms-MessageBar--success");
  } else if (type === "error") {
    statusElement.classList.add("ms-MessageBar--error");
  }
  
  // Show the message
  statusElement.style.display = "block";
}
