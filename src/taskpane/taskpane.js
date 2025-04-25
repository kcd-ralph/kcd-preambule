let reports = [];
let clientLogoBase64 = null;

Office.onReady(() => {
  loadReports();

  document.getElementById("client-logo").addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (file) {
      try {
        clientLogoBase64 = await convertBlobToBase64(file);
        setStatus("✅ Logo client chargé.");
      } catch (e) {
        setError("Erreur lors du chargement du logo client.");
      }
    } else {
      clientLogoBase64 = null;
    }
  });

  document.getElementById("insert-btn").onclick = async () => {
    const select = document.getElementById("report-select");
    const selectedIndex = select.value;
    const clientName = document.getElementById("client-name").value.trim();

    if (!selectedIndex) {
      setError("Veuillez sélectionner un rapport.");
      return;
    }

    if (!clientName) {
      setError("Veuillez entrer le nom du client.");
      return;
    }

    const report = reports[selectedIndex];

    try {
      await insertReport(report.title, clientName, report.vc);
      setStatus(`✅ Rapport inséré : ${report.title}`);
    } catch (err) {
      console.error("Insert failed:", err);
      setError(`⚠️ Échec de l'insertion : ${err.message || err}`);
    }
  };
});

async function loadReports() {
  try {
    const response = await fetch("https://kcd-ralph.github.io/kcd-preambule/src/titles.json");
    reports = await response.json();

    const select = document.getElementById("report-select");
    select.innerHTML = '<option value="">-- Sélectionnez un rapport --</option>';

    reports.forEach((report, index) => {
      const option = document.createElement("option");
      option.value = index;
      option.textContent = report.title;
      select.appendChild(option);
    });

    setStatus("✅ Rapports chargés.");
  } catch (err) {
    setError("Impossible de charger le fichier titles.json.");
  }
}

async function insertReport(title, clientName, includeVC) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.clear();

    // Karman logo
    const karmanLogo = await getImageBase64("assets/karman.png");
    const karmanPara = body.insertParagraph("", Word.InsertLocation.end);
    karmanPara.alignment = Word.Alignment.centered;
    karmanPara.insertInlinePictureFromBase64(karmanLogo, Word.InsertLocation.end);

    // Report title and client
    const titlePara = body.insertParagraph(title, Word.InsertLocation.end);
    titlePara.alignment = Word.Alignment.centered;
    titlePara.font.set({ name: "Calibri", size: 26, bold: true, color: "#11255E" });

    const clientPara = body.insertParagraph(`Client : ${clientName}`, Word.InsertLocation.end);
    clientPara.alignment = Word.Alignment.centered;
    clientPara.font.set({ name: "Calibri", size: 16, color: "#11255E" });

    // Spacer
    for (let i = 0; i < 5; i++) body.insertParagraph("", Word.InsertLocation.end);

    // Optional client logo
    if (clientLogoBase64) {
      const logoPara = body.insertParagraph("", Word.InsertLocation.end);
      logoPara.alignment = Word.Alignment.centered;
      logoPara.insertInlinePictureFromBase64(clientLogoBase64, Word.InsertLocation.end);
    }

    // Legal notice
    body.insertParagraph("", Word.InsertLocation.end);
    body.insertParagraph("AVISLÉGAL", Word.InsertLocation.end).font.set({
      name: "Times New Roman",
      italic: true,
      size: 8,
      color: "#000000",
      bold: false,
    });

    body
      .insertParagraph(
        "Ce document est destiné uniquement à l'usage de la personne ou de l'entité à laquelle il est adressé et peut contenir des informations privilé-giées, confidentielles et exemptes de divulgation en vertu de la loi applicable. Si le lecteur de cet avis légal n'est pas le destinataire prévu, nous vous informons par la présente que toute diffusion, distribution ou copie de ce document est strictement interdite. Si vous avez reçu ce document par erreur, veuillez-nous en informer immédiatement par téléphone et nous retourner l'original à l'adresse postale ci-dessous.",
        Word.InsertLocation.end
      )
      .font.set({ name: "Times New Roman", italic: true, size: 8, bold: false, color: "#000000",});

    // Flag logo
    const flagLogo = await getImageBase64("assets/flag.png");
    const flagPara = body.insertParagraph("", Word.InsertLocation.end);
    flagPara.alignment = Word.Alignment.centered;
    flagPara.insertInlinePictureFromBase64(flagLogo, Word.InsertLocation.end);

    const address = ["1 allée des écureuils", "69380 Lissieu, France", "+33 (0)4 72 54 88 58"];
    address.forEach((line) => {
      const p = body.insertParagraph(line, Word.InsertLocation.end);
      p.alignment = Word.Alignment.centered;
      p.font.set({ name: "Calibri", size: 9, bold: false, color: "#000000", });
    });

    if (includeVC) {
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

      // === Table 1: Contrôle des documents ===
      body.insertParagraph("Contrôle des documents", Word.InsertLocation.end).font.set({
        bold: true,
        size: 16,
        color: "#11255E",
      });

      const table1 = body.insertTable(3, 4, Word.InsertLocation.end, [
        ["", "Nom", "Fonction", "Date"],
        ["Écrit par", "XYZ", "XYZ", "XYZ"],
        ["Validé par", "XYZ", "XYZ", "XYZ"],
      ]);

      table1.load("rows/items/cells");
      await context.sync();

      body.insertParagraph("", Word.InsertLocation.end);

      const rows1 = table1.rows;
      rows1.load("items");
      await context.sync();

      const headerRow1 = table1.rows.items[0];
      styleHeaderRow(headerRow1);

      for (let i = 1; i < rows1.items.length; i++) {
        rows1.items[1].setCellPadding(Word.CellPaddingLocation.top, 5);
        rows1.items[1].setCellPadding(Word.CellPaddingLocation.bottom, 5);

        rows1.items[2].setCellPadding(Word.CellPaddingLocation.top, 5);
        rows1.items[2].setCellPadding(Word.CellPaddingLocation.bottom, 5);

        const row1 = rows1.items[i];
        row1.font.size = 12;
        row1.verticalAlignment = Word.VerticalAlignment.center;

        row1.cells.load("items");
        await context.sync();

        const cells1 = row1.cells.items;

        // Bold only the first column (column 0) in rows 1–5
        if (i >= 1 && i <= 3 && cells1[0]) {
          cells1[0].shadingColor = "#11255E";
        }

        // Alignment setup
        cells1.forEach((cell) => {
          cell.horizontalAlignment = Word.Alignment.centered;
        });
      }

      await context.sync();

      // === Table 2: Contrôle de version ===
      body.insertParagraph("Contrôle de version", Word.InsertLocation.end).font.set({
        bold: true,
        size: 16,
        color: "#11255E",
      });

      const table2 = body.insertTable(2, 4, Word.InsertLocation.end, [
        ["Numéro de version", "Date", "Auteur", "Nature du changement"],
        ["1.0", "XYZ", "XYZ", "XYZ"],
      ]);

      table2.load("rows/items/cells");
      await context.sync();

      body.insertParagraph("", Word.InsertLocation.end);

      const rows2 = table2.rows;
      rows1.load("items");
      await context.sync();

      const headerRow2 = table2.rows.items[0];
      styleHeaderRow(headerRow2);

      for (let i = 1; i < rows2.items.length; i++) {
        rows2.items[1].setCellPadding(Word.CellPaddingLocation.top, 5);
        rows2.items[1].setCellPadding(Word.CellPaddingLocation.bottom, 5);

        const row2 = rows2.items[i];
        row2.font.size = 12;
        row2.verticalAlignment = Word.VerticalAlignment.center;

        row2.cells.load("items");
        await context.sync();

        const cells2 = row2.cells.items;

        // Alignment setup
        cells2.forEach((cell) => {
          cell.horizontalAlignment = Word.Alignment.centered;
        });
      }

      await context.sync();

      // === Table 3: Distribution ===
      body.insertParagraph("Distribution", Word.InsertLocation.end).font.set({
        bold: true,
        size: 16,
        color: "#11255E",
      });

      const table3 = body.insertTable(2, 2, Word.InsertLocation.end, [
        ["Nom", "Fonction"],
        ["XYZ", "XYZ"],
      ]);

      table3.load("rows/items/cells");
      await context.sync();

      const rows3 = table3.rows;
      rows1.load("items");
      await context.sync();

      const headerRow3 = table3.rows.items[0];
      styleHeaderRow(headerRow3);

      for (let i = 1; i < rows3.items.length; i++) {
        rows3.items[1].setCellPadding(Word.CellPaddingLocation.top, 5);
        rows3.items[1].setCellPadding(Word.CellPaddingLocation.bottom, 5);

        const row3 = rows3.items[i];
        row3.font.size = 12;
        row3.verticalAlignment = Word.VerticalAlignment.center;

        row3.cells.load("items");
        await context.sync();

        const cells3 = row3.cells.items;

        // Alignment setup
        cells3.forEach((cell) => {
          cell.horizontalAlignment = Word.Alignment.centered;
        });
      }

      await context.sync();
    }

    await context.sync();
  });
}

async function getImageBase64(url) {
  const response = await fetch(url);
  const blob = await response.blob();
  return await convertBlobToBase64(blob);
}

function convertBlobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function styleHeaderRow(row) {
  row.shadingColor = "#11255E";
  row.setCellPadding(Word.CellPaddingLocation.top, 5);
  row.setCellPadding(Word.CellPaddingLocation.bottom, 5);
  row.font.size = 12;
  row.font.bold = true;
  row.horizontalAlignment = Word.Alignment.centered;
  row.verticalAlignment = Word.VerticalAlignment.center;
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
  document.getElementById("error").textContent = "";
}

function setError(message) {
  document.getElementById("error").textContent = message;
  document.getElementById("status").textContent = "";
}
