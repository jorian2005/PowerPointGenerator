const express = require("express");
const pptxgen = require("pptxgenjs");
const path = require("path");

const app = express();
const port = 3000;

app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));

const fetch = (...args) => import("node-fetch").then(({ default: fetch }) => fetch(...args));

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

app.get("/", (req, res) => {
  res.render("index", { dominee: "", ouderling: "", organist: "", liederen: "", bestandsnaam: "", datum: "", dienst: "" });
});

app.post("/generate", async (req, res) => {
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
  }

  const liederen = req.body.liederen.split("\n").filter(Boolean);
  const datum = new Date(req.body.datum).toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
  const dienst = req.body.dienst;

  const collectes = [
    { name: req.body.collecte1, image: req.body.collecte1Image },
    { name: req.body.collecte2, image: req.body.collecte2Image },
    { name: req.body.collecte3, image: req.body.collecte3Image },
  ];

  const pptx = new pptxgen();
  const slide = pptx.addSlide();
  const addTextWithImage = async (text, imagePath, x, y, w, h) => {
    if (imagePath.startsWith("http")) {
      const imageBuffer = await urlToImage(imagePath);
      slide.addImage({ data: imageBuffer, x, y: y + 0.5, w, h });
    } else {
      slide.addImage({ path: imagePath, x, y: y + 0.5, w, h });
    }
    slide.addText(text, { x, y, fontSize: 15, bold: true });
  };

  const xPosition = 1;
  const yPosition = 3;

  // Aankondiging van de dienst
  addTextWithImage(`Welkom in de ${dienst} van ${datum} `, "", xPosition, 1, 6, 1);
  slide.addText("", { x: xPosition, y: yPosition + 4, w: 6, h: 1, align: "center" });

  // Medewerkenden op Aankondigingsslide
  let xPos = xPosition;

  // dominee
  meewerkenden.dominee.forEach((persoon) => {
    const imagePath = path.resolve(__dirname, "public", "dominee.jpg");
    addTextWithImage(`Dominee: ${persoon}`, imagePath, xPos, yPosition, 1, 1);
    xPos += 3;
  });

  // ouderling
  meewerkenden.ouderling.forEach((persoon) => {
    const imagePath = path.resolve(__dirname, "public", "ouderling.jpg");
    addTextWithImage(`Ouderling: ${persoon}`, imagePath, xPos, yPosition, 1, 1);
    xPos += 3;
  });

  // organist
  meewerkenden.organist.forEach((persoon) => {
    const imagePath = path.resolve(__dirname, "public", "organist.jpg");
    addTextWithImage(`Organist: ${persoon}`, imagePath, xPos, yPosition, 1, 1);
    xPos += 3;
  });

  // Collectes toevoegen
  for (let i = 0; i < collectes.length; i++) {
    const collecte = collectes[i];
    const { name, image } = collecte;
    if (name && image) {
      let imagePath = image.trim();
      if (imagePath.startsWith("http")) {
        const imageBuffer = await urlToImage(imagePath);
        slide.addImage({ data: imageBuffer, x: xPosition, y: yPosition + 4 + i, w: 1, h: 1 });
      } else {
        imagePath = path.resolve(__dirname, "public", imagePath);
        slide.addImage({ path: imagePath, x: xPosition, y: yPosition + 4 + i, w: 1, h: 1 });
      }
      slide.addText(`Collecte ${i + 1}: ${name}`, { x: xPosition + 1, y: yPosition + 4 + i, w: 5, h: 1 });
    }
  }

  // Save the PowerPoint file
  const filename = req.body.bestandsnaam || "meewerkenden_en_liederen.pptx";
  pptx.writeFile({ fileName: `public/${filename}` }, (error) => {
    if (error) {
      console.error(error);
      res.send("Er is een fout opgetreden bij het genereren van de PowerPoint.");
    } else {
      console.log("PowerPoint-bestand gegenereerd!");
      res.download(`public/${filename}`);
    }
  });
});

app.listen(port, () => {
  console.log(`Server gestart op http://localhost:${port}`);
});
