/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import Papa from 'papaparse';
import {base64Image} from "../../base64Image";


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("gen-slips").onclick = () => tryCatch(GenerateSlips);
    document.getElementById("test").onclick = () => tryCatch(insertSimpleText);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});


function stream(abbr){
  const streams = {
    'Processing': 'Jr. Game Developer II - CVA (C500)',
    'Rob Eng I': 'Robotics Engineer I (R400)',
    'Scratch': 'Jr. Game Developer I - Scratch (C200)',
    'Mov. Models': 'Moving Models (R300)',
    'Rob Eng II': 'Robotics Engineer II (R600)',
    'Minecraft': 'Minecraft Adventures (C350)',
    'Jr. Rob Eng': 'Moving Models (R300)',
    'Python': 'Python Developer (C800)',
    'Unity': 'Sr. Game Developer - Unity (C700)',
    'Web Dev': 'Web Developer (C620)',
    'Sr. Rob Eng I': 'Sr Robotics Engineer I - VEX (R800)',
    'Roblox': 'Roblox Adventures (C650)'
    }
  return streams[abbr];
}

function password(name){
  if (name.length >= 8){
    return name;
  }else if (name.length < 5){
    return name + "12345678".slice(0, 8 - name.length);
  }else {
    return name + "123";
  }
}

const ZebraHeading = {
  name: "Arial",
  size: 24,
  bold: true,
  color: "#54AFE2"
}
const ZebraBold = {
  name: "Arial",
  size: 20,
  bold: true,
  color: "black"
}
const ZebraPlain = {
  name: "Arial",
  size: 20,
  bold: false,
  color: "black"
}


async function GenerateSlips() {

  await Word.run(async (context) => {
      // Queue commands to insert a paragraph into the document.
      const docBody = context.document.body;
      
      const selectedFile = document.getElementById("students").files[0];
      
      Papa.parse(selectedFile, {
        complete: function(results) {
          const students = results["data"];
          console.log(results);



          for (const student of students){
            // blank space between slips
            docBody.insertParagraph("", Word.InsertLocation.start)
            docBody.insertParagraph("", Word.InsertLocation.start)
            docBody.insertParagraph("", Word.InsertLocation.start)
            docBody.insertParagraph("", Word.InsertLocation.start)
            docBody.insertParagraph("", Word.InsertLocation.start)
            
            // course
            const course = docBody.insertText(stream(student["Stream (Abbr)"]), "Start");
            docBody.insertText("Course:         ", Word.InsertLocation.start).font.set(ZebraBold);
            
            //password
            docBody.insertParagraph("", "Start");
            const pass = docBody.insertText(password(student["Student Name"].split(" ")[0]), "Start")
            docBody.insertText("Password:    ", Word.InsertLocation.start).font.set(ZebraBold);
            
            //email
            docBody.insertParagraph("", "Start");
            docBody.insertText(student["Student ID"] + "@zebrarobotics.com", Word.InsertLocation.start).font.set(ZebraPlain);
            docBody.insertText("Email:           ", Word.InsertLocation.start).font.set(ZebraBold);

            //name
            var paragraph = docBody.insertParagraph("Student:       " + student["Student Name"], Word.InsertLocation.start);
            paragraph.font.set(ZebraBold);
            //heading
            docBody.insertParagraph("", "Start");
            docBody.insertParagraph("", "Start");
            docBody.insertInlinePictureFromBase64(base64Image, "Start");
            
            // title.font.set(ZebraHeading);
            course.font.set(ZebraPlain);
            pass.font.set(ZebraPlain);


          }
          
        }, 
        header:true
      });
      
     
      
      await context.sync();
  });

  


// DOES NOT WORK - TODO: figure out how to set word document format and paragraph spacing
// Word.run(async (context) => {
//     // Get all paragraphs in the document
//     const paragraphs = context.document.body.paragraphs;

//     // Load the paragraphs and their paragraphFormat properties
//     paragraphs.load("items");

//     await context.sync(); // Sync to load the paragraphs

//     // Loop through each paragraph and load the paragraphFormat
//     paragraphs.items.forEach(paragraph => {
//       console.log(paragraph.style.paragraphFormat);
//       paragraph.style.paragraphFormat.load(); // Load the paragraphFormat object
      
//     });

//     // Loop through each paragraph and set line spacing
//     paragraphs.items.forEach(paragraph => {
//         paragraph.style.paragraphFormat.lineSpacing = 1; // Set line spacing to double spacing
//     });

//     await context.sync(); // Sync changes to the document
// })
// .catch(function (error) {
//     console.log("Error: " + error);
// });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}

