import express from "express";
import pptxgen from "pptxgenjs";
import path from "path";
import fs from "fs";
import fetch from "node-fetch";
import { fileURLToPath } from 'url';

const app = express();
const port = 3000;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));

const urlToImage = async (url) => {
  try {
    const response = await fetch(url);
    const buffer = await response.buffer();
    return buffer;
  } catch (error) {
    console.error("Error fetching image:", error);
    return null;
  }
};

const addTextWithImage = async (slide, text, imagePath, x, y, w, h) => {
  try {
    let imageOptions = {};
    if (imagePath.startsWith("http")) {
      const imageBuffer = await urlToImage(imagePath);
      imageOptions.data = imageBuffer;
    } else {
      let absoluteImagePath = imagePath;
      if (!imagePath.startsWith("http")) {
        const publicImagePath = `public/${imagePath}`;
        absoluteImagePath = path.join(__dirname, publicImagePath);
      }

      // Kijk of het bestand bestaat en of het een bestand is
      if (fs.existsSync(absoluteImagePath) && fs.statSync(absoluteImagePath).isFile()) {
        imageOptions.path = absoluteImagePath;
        slide.addImage(Object.assign(imageOptions, { x, y: y + 0.2, w, h }));
      } else {
        console.error("Image file does not exist or is not a file:", absoluteImagePath);
      }
    }
    slide.addText(text, { x: x + 0.005, y, fontSize: 15, bold: true });
  } catch (error) {
    console.error("Error adding text with image:", error);
  }
};

app.get("/", (req, res) => {
  const indexPath = path.join(__dirname, "public", "index.html");
  res.sendFile(indexPath);
});

app.post("/generate", async (req, res) => {
  // changeImage();
  let meewerkenden = {
    dominee: [],
    ouderling: [],
    organist: [],
  };

  if (req.body.dominee) {
    meewerkenden.dominee = req.body.dominee.split("\n").filter(Boolean);
  }
  if (req.body.ouderling) {
    meewerkenden.ouderling = req.body.ouderling.split("\n").filter(Boolean);
  }
  if (req.body.organist) {
    meewerkenden.organist = req.body.organist.split("\n").filter(Boolean);
  };

  const liederen = req.body.liederen.split("\n").filter(Boolean);
  const datum = new Date(req.body.datum).toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
  const dienst = req.body.dienst;

  const collectes = [
    { name: req.body.collecte1, image: req.body.collecte1Image },
    { name: req.body.collecte2, image: req.body.collecte2Image },
    { name: req.body.collecte3, image: req.body.collecte3Image },
  ];

  const pptx = new pptxgen();

  // function changeImage() {
  //   var selectBox = document.getElementById("kerkgebouw");
  //   var selectedKerk = selectBox.value;

  //   var imagePath = "";
  //   if (selectedKerk === "Dorpskerk") {
  //     imagePath = "public/Dorpskerk.png";
  //   } else if (selectedKerk === "Oenenburgkerk") {
  //     imagePath = "public/Oenenburgkerk.png";
  //   }

  //   var kerkAfbeelding = document.getElementById("kerkAfbeelding");
  //   kerkAfbeelding.src = imagePath;
  // }

  // Maak de eerste dia aan
  const welkomSlide = pptx.addSlide();

  // Positie van de tekst en afbeelding
  const xPosition = 1;
  const yPosition = 1;
  const xPos = 1;

  // Tekst en afbeelding toevoegen aan de welkom dia
  welkomSlide.addText(`Welkom in de ${dienst} van ${datum} `, { x: 0, y: 0.5, w: 10, h: 1, align: "center", fontSize: 30, bold: true });
  welkomSlide.addImage({ path: "public/dorpskerk.png", x: xPosition + 1, y: yPosition + 0.5, w: 6, h: 2 });
  // Meewerkenden
  addTextWithImage(welkomSlide, `Dominee: ${meewerkenden.dominee}`, "dominee.jpg", xPos, 4, 1, 1);
  addTextWithImage(welkomSlide, `Ouderling: ${meewerkenden.ouderling}`, "ouderling.jpg", xPos + 3, 4, 1, 1);
  addTextWithImage(welkomSlide, `Organist: ${meewerkenden.organist}`, "organist.jpg", xPos + 6, 4, 1, 1);

  // Maak de tweede dia aan voor collecte
  const collectesSlide = pptx.addSlide();

  // Titel toevoegen aan de collectes dia
  collectesSlide.addText("Collectes", { x: 0.5, y: 0.5, w: 8, h: 1, align: "center", fontSize: 30, bold: true });

  // Afbeeldingen en tekst toevoegen aan de collectes dia
  for (let i = 0; i < collectes.length; i++) {
    const collecte = collectes[i];
    const { name, image } = collecte;
    if (name && image) {
      let imagePath = image.trim();
      if (imagePath.startsWith("http")) {
        const imageBuffer = await urlToImage(imagePath);
        collectesSlide.addImage({ data: imageBuffer, x: xPosition, y: yPosition + 0.5 + i, w: 1, h: 1 });
      } else {
        imagePath = path.join(__dirname, "public", imagePath);
        collectesSlide.addImage({ path: imagePath, x: xPosition, y: yPosition + 0.5 + i, w: 1, h: 1 });
      }
      collectesSlide.addText(`Collecte ${i + 1}: ${name}`, { x: xPosition + 1, y: yPosition + 0.5 + i, w: 8, h: 1 });
    }
  }

  const filename = req.body.bestandsnaam || "meewerkenden_en_liederen.pptx";
  const filepath = path.join(__dirname, "public", filename);

  try {
    const buffer = await pptx.write("nodebuffer");
    console.log("Generated PowerPoint Buffer:", buffer); // Log the buffer

    // Bestand schrijven naar de server
    fs.writeFile(filepath, buffer, (error) => {
      if (error) {
        console.error("Error writing PowerPoint file:", error);
        res.send("An error occurred while generating the PowerPoint.");
      } else {
        console.log("PowerPoint file generated!");
        // Stuur het bestand naar de client
        res.download(filepath, filename, (downloadError) => {
          if (downloadError) {
            console.error("Error downloading the file:", downloadError);
          } else {
            console.log("File downloaded successfully!");
          }
          // Verwijder het bestand van de server
          fs.unlink(filepath, (unlinkError) => {
            if (unlinkError) {
              console.error("Error removing the file:", unlinkError);
            } else {
              console.log("File removed successfully!");
            }
          });
        });
      }
    });
  } catch (error) {
    console.error("Error generating PowerPoint:", error);
    res.send("An error occurred while generating the PowerPoint.");
  }
});

app.listen(port, () => {
  console.log(`Server gestart op http://localhost:${port}`);
});