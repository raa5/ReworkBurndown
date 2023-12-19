let myChart; // Declare myChart variable outside the handleFile function to make it accessible globally

document.getElementById("fileInput").addEventListener("change", handleFile);

function handleFile(event) {
  // Clear the previous chart and history
  if (myChart) {
    myChart.destroy();
    document
      .getElementById("myChart")
      .getContext("2d")
      .clearRect(
        0,
        0,
        document.getElementById("myChart").width,
        document.getElementById("myChart").height
      );
  }

  const fileInput = event.target;
  const file = fileInput.files[0];

  if (file) {
    const reader = new FileReader();

    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Clear the sheet links container
      const sheetLinksContainer = document.getElementById("sheetLinks");
      sheetLinksContainer.innerHTML = "";

      // Create clickable links for each sheet
      workbook.SheetNames.forEach((sheetName) => {
        const sheetLink = document.createElement("div");
        sheetLink.classList.add("retro-sheet-link");
        sheetLink.textContent = sheetName;
        sheetLink.addEventListener("click", () =>
          processSheet(workbook, sheetName)
        );
        sheetLinksContainer.appendChild(sheetLink);
      });
    };

    reader.readAsArrayBuffer(file);
  }
}

function processSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];

  // Clear the previous chart
  if (myChart) {
    myChart.destroy();
    document
      .getElementById("myChart")
      .getContext("2d")
      .clearRect(
        0,
        0,
        document.getElementById("myChart").width,
        document.getElementById("myChart").height
      );
  }

  const lastRowIndex = XLSX.utils.decode_range(sheet["!ref"]).e.r;
  let duaFailsRowIndex = null;
  const labels = [];
  const values = [];

  for (let i = 5; i <= lastRowIndex; i++) {
    const dateCellAddress = "A" + (i + 1);
    const dateValue = sheet[dateCellAddress] ? sheet[dateCellAddress].w : "";

    const valueCellAddress = "B" + (i + 1);
    const cellValue = sheet[valueCellAddress] ? sheet[valueCellAddress].v : "";
    const floatValue = parseFloat(cellValue);

    if (dateValue && !isNaN(floatValue)) {
      // Include date and value in the arrays for chart
      labels.push(dateValue);
      values.push(floatValue);
    }

    // Check if the current row contains "DUA fails"
    if (cellValue === "DUA fails") {
      duaFailsRowIndex = i;
      break;
    }
  }

  if (duaFailsRowIndex !== null) {
    // Create an animated line chart with a retro theme using Chart.js
    const ctx = document.getElementById("myChart").getContext("2d");
    myChart = new Chart(ctx, {
      type: "line",
      data: {
        labels: labels,
        datasets: [
          {
            label: "Values",
            data: values,
            borderColor: "rgba(255, 193, 7, 1)",
            backgroundColor: "rgba(255, 193, 7, 0.2)",
            borderWidth: 2,
            pointRadius: 4,
            pointBackgroundColor: "rgba(255, 193, 7, 1)",
            pointBorderColor: "rgba(255, 193, 7, 1)",
            pointHoverRadius: 6,
          },
        ],
      },
      options: {
        scales: {
          x: [
            {
              type: "category",
              labels: labels,
              position: "bottom",
              ticks: {
                fontColor: "#0de8f04d",
              },
            },
          ],
          y: [
            {
              type: "linear",
              position: "left",
              ticks: {
                fontColor: "#0de8f04d",
              },
            },
          ],
        },
        legend: {
          display: true,
          labels: {
            fontColor: "#fff",
          },
        },
        title: {
          display: true,
          text: `Retro Excel Chart - ${sheetName}`,
          fontColor: "#fff",
        },
        animation: {
          duration: 2000,
          easing: "easeInOutQuart",
        },
      },
    });

    // Update chart on window resize
    window.addEventListener("resize", function () {
      myChart.update();
    });
  } else {
    console.log('String "DUA fails" not found in column B.');
  }
}
