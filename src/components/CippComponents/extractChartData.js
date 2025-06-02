/**
 * Enhanced method to extract chart data from a chart element
 * @param {HTMLElement} card - The card element containing the chart
 * @param {HTMLElement} chartElement - The chart element
 * @param {string} chartType - The detected chart type
 * @returns {Object} - The extracted chart data
 */
const extractChartData = (card, chartElement, chartType) => {
  const chartData = {
    type: chartType,
    labels: [],
    values: [],
    totalValue: 0,
  };

  // Find total value from typography elements
  const totalValueElement = card.querySelector(".MuiTypography-h5");
  if (totalValueElement && totalValueElement.nextElementSibling) {
    chartData.totalValue = totalValueElement.nextElementSibling.textContent.trim();
  }

  // Extract data from DOM - look for data attributes
  try {
    const chartContainer = card.querySelector("[data-chart-data]");
    if (chartContainer) {
      const chartDataAttr = chartContainer.getAttribute("data-chart-data");
      if (chartDataAttr) {
        try {
          const parsedData = JSON.parse(chartDataAttr);
          if (parsedData.labels && Array.isArray(parsedData.labels)) {
            chartData.labels = parsedData.labels;
          }
          if (parsedData.series && Array.isArray(parsedData.series)) {
            chartData.values = parsedData.series.map(String);
          }

          // If we successfully extracted data, return it
          if (chartData.labels.length > 0) {
            return chartData;
          }
        } catch (err) {
          console.warn("Failed to parse chart data attribute:", err);
        }
      }
    }
  } catch (err) {
    console.warn("Error extracting chart data from attributes:", err);
  }

  // Extract labels and values from the legend items below the chart
  const dataItems = card.querySelectorAll(".MuiStack-root .MuiStack-root");
  if (dataItems.length > 0) {
    // Clear any previous data if we found items in the legend
    chartData.labels = [];
    chartData.values = [];

    dataItems.forEach((item) => {
      const label = item.querySelector(".MuiTypography-body2");
      const value = label ? item.querySelectorAll(".MuiTypography-body2")[1] : null;

      if (label && value) {
        chartData.labels.push(label.textContent.trim());
        chartData.values.push(value.textContent.trim());
      }
    });

    // If we successfully extracted data, return it
    if (chartData.labels.length > 0) {
      return chartData;
    }
  }

  // If all else fails, try to extract data from chart SVG elements
  const allTexts = Array.from(card.querySelectorAll("text.apexcharts-text"));
  const possibleLabels = allTexts
    .filter(
      (textEl) =>
        textEl.classList.contains("apexcharts-datalabel-label") ||
        textEl.classList.contains("apexcharts-xaxis-label")
    )
    .map((textEl) => textEl.textContent.trim())
    .filter((text) => text.length > 0);

  if (possibleLabels.length > 0) {
    chartData.labels = possibleLabels;

    // Try to find corresponding values
    const possibleValues = allTexts
      .filter((textEl) => textEl.classList.contains("apexcharts-datalabel-value"))
      .map((textEl) => textEl.textContent.trim());

    if (possibleValues.length > 0) {
      chartData.values = possibleValues;
    }
  }

  return chartData;
};

export default extractChartData;
