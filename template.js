console.log("template.js wurde geladen");
function generate() {
  let pptx = new PptxGenJS();

  pptx.layout = "LAYOUT_16x9";

  const COLOR_TEXT = "222222";
  const COLOR_SUB  = "555555";
  const COLOR_RED  = "CC3333";
  const COLOR_LINE = "DDDDDD";
  const FONT_MAIN  = "Arial";

  function addNotebookLines(slide) {
    const lineCount = 14;
    for (let i = 0; i < lineCount; i++) {
      slide.addShape(pptx.ShapeType.LINE, {
        x: 0.5,
        y: 0.8 + i * 0.4,
        w: 9,
        h: 0,
        line: { color: COLOR_LINE, width: 0.5 }
      });
    }
  }

  function addGridItems(slide) {
    const cols = 4;
    const rows = 5;
    const startX = 0.6;
    const startY = 1.2;
    const cellW = 2.2;
    const cellH = 1.4;

    let index = 1;

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const x = startX + c * cellW;
        const y = startY + r * cellH;

        slide.addShape(pptx.ShapeType.ELLIPSE, {
          x: x,
          y: y,
          w: 0.45,
          h: 0.45,
          line: { color: COLOR_RED, width: 1 },
          fill: { color: "FFFFFF" }
        });

        slide.addText(`Item ${index}`, {
          x: x + 0.55,
          y: y,
          w: cellW - 0.6,
          h: 0.4,
          fontFace: FONT_MAIN,
          fontSize: 12,
          bold: true,
          color: COLOR_TEXT
        });

        slide.addText("1–2 Zeilen Beschreibung", {
          x: x + 0.55,
          y: y + 0.42,
          w: cellW - 0.6,
          h: 0.8,
          fontFace: FONT_MAIN,
          fontSize: 10,
          color: COLOR_SUB
        });

        index++;
      }
    }
  }

  let slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  addNotebookLines(slide);

  slide.addText("20 Lektionen aus 4 Jahren KI-Powernutzung", {
    x: 0.6,
    y: 0.3,
    w: 9,
    h: 0.5,
    fontFace: FONT_MAIN,
    fontSize: 22,
    bold: true,
    color: COLOR_TEXT
  });

  slide.addText("Minimalist Flat Infographic – Notebook Style", {
    x: 0.6,
    y: 0.75,
    w: 9,
    h: 0.3,
    fontFace: FONT_MAIN,
    fontSize: 12,
    color: COLOR_SUB
  });

  addGridItems(slide);

  slide.addText("Was war DEINE größte Lektion?", {
    x: 0.6,
    y: 8.0,
    w: 9,
    h: 0.5,
    fontFace: FONT_MAIN,
    fontSize: 16,
    bold: true,
    color: COLOR_RED
  });

  pptx.writeFile({ fileName: "Minimalist_Flat_Infographic_Template.pptx" });
}

