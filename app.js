import express from "express";
import pptxgen from "pptxgenjs";
import path from "path";
import fs from "fs";
import fetch from "node-fetch";
import { fileURLToPath } from 'url';
import { publicDecrypt } from "crypto";

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

      // Check of het bestand bestaat en of het een bestand is
      if (fs.existsSync(absoluteImagePath) && fs.statSync(absoluteImagePath).isFile()) {
        imageOptions.path = absoluteImagePath;
        slide.addImage(Object.assign(imageOptions, { x, y: y + 0.2, w, h }));
      } else {
        console.error("Image file does not exist or is not a file:", absoluteImagePath);
      }
    }
    slide.addText(text, { x, y, fontSize: 15, bold: true });
  } catch (error) {
    console.error("Error adding text with image:", error);
  }
};

app.post("/generate", async (req, res) => {
  
});

app.listen(port, () => {
  console.log(`Server gestart op http://localhost:${port}`);
});