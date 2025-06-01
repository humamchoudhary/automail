function loadFonts(fontFiles, fontFolder) {
  fontFiles.forEach((fileName) => {
    let fontWeight = "400";
    let fontStyle = "normal";

    if (fileName.includes("Black")) fontWeight = "900";
    if (fileName.includes("Bold")) fontWeight = "700";
    if (fileName.includes("Semi")) fontWeight = "800";
    if (fileName.includes("Light")) fontWeight = "300";
    if (fileName.includes("Thin")) fontWeight = "100";
    if (fileName.includes("Italic")) fontStyle = "italic";

    const style = document.createElement("style");
    style.innerHTML = `
            @font-face {
                font-family: 'Proxima Nova';
                src: url('${fontFolder}/${fileName}') format('${getFormat(fileName)}');
                font-weight: ${fontWeight};
                font-style: ${fontStyle};
                font-display: swap;
            }
        `;
    document.head.appendChild(style);
  });
}

function getFormat(fileName) {
  if (fileName.endsWith(".ttf")) return "truetype";
  if (fileName.endsWith(".otf")) return "opentype";
  if (fileName.endsWith(".woff")) return "woff";
  if (fileName.endsWith(".woff2")) return "woff2";
  return "truetype"; // default
}

// loadFonts(fontFiles, fontFolder);
