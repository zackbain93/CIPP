import React, { useState, useEffect, useCallback } from "react";
import {
  Document,
  Page,
  Text,
  View,
  StyleSheet,
  pdf,
  Font,
  Image,
  PDFViewer,
} from "@react-pdf/renderer";
import { Dialog, DialogContent, DialogTitle, IconButton } from "@mui/material";
import CloseIcon from "@mui/icons-material/Close";
import { saveAs } from "file-saver";
import toast from "react-hot-toast";
import extractChartData from "./extractChartData";

// Define styles for PDF document
const styles = StyleSheet.create({
  page: {
    flexDirection: "column",
    backgroundColor: "#ffffff",
    padding: 20,
    fontFamily: "Helvetica",
  },
  header: {
    marginBottom: 20,
    borderBottomWidth: 1,
    borderBottomColor: "#e0e0e0",
    borderBottomStyle: "solid",
    paddingBottom: 10,
  },
  title: {
    fontSize: 18,
    fontWeight: "bold",
    marginBottom: 5,
    color: "#333333",
  },
  subtitle: {
    fontSize: 12,
    color: "#666666",
    marginBottom: 10,
  },
  timestamp: {
    fontSize: 10,
    color: "#999999",
    marginBottom: 5,
  },
  content: {
    marginTop: 10,
    fontSize: 10,
    lineHeight: 1.5,
  },
  paragraph: {
    fontSize: 10,
    marginBottom: 5,
    lineHeight: 1.5,
  },
  heading: {
    fontWeight: "bold",
    marginTop: 12,
    marginBottom: 6,
    color: "#333333",
  },
  tableContainer: {
    marginTop: 15,
    marginBottom: 15,
  },
  tableTitle: {
    fontSize: 12,
    fontWeight: "bold",
    marginBottom: 5,
    color: "#333333",
  },
  table: {
    display: "table",
    width: "100%",
    borderStyle: "solid",
    borderWidth: 1,
    borderColor: "#e0e0e0",
  },
  tableRow: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderBottomColor: "#e0e0e0",
    borderBottomStyle: "solid",
  },
  tableRowEven: {
    backgroundColor: "#ffffff",
  },
  tableRowOdd: {
    backgroundColor: "#f9f9f9",
  },
  tableRowHeader: {
    flexDirection: "row",
    backgroundColor: "#f5f5f5",
    borderBottomWidth: 1,
    borderBottomColor: "#e0e0e0",
    borderBottomStyle: "solid",
  },
  tableCell: {
    padding: 5,
    fontSize: 8,
    minWidth: 60,
    maxWidth: 200,
    wordBreak: "break-word",
    // Removed flex for better PDF layout
  },
  tableCellHeader: {
    padding: 5,
    fontSize: 8,
    fontWeight: "bold",
    minWidth: 60,
    maxWidth: 200,
    wordBreak: "break-word",
    // Removed flex for better PDF layout
  },
  noData: {
    marginTop: 30,
    fontSize: 14,
    textAlign: "center",
    color: "#999999",
  },
  image: {
    width: "100%",
    marginVertical: 10,
  },
  footer: {
    position: "absolute",
    bottom: 20,
    left: 20,
    right: 20,
    fontSize: 8,
    color: "#999999",
    textAlign: "center",
    borderTopWidth: 1,
    borderTopColor: "#e0e0e0",
    borderTopStyle: "solid",
    paddingTop: 5,
  },
  pageNumber: {
    position: "absolute",
    bottom: 10,
    right: 20,
    fontSize: 8,
    color: "#999999",
  },
  card: {
    marginBottom: 15,
    padding: 10,
    borderWidth: 1,
    borderColor: "#e0e0e0",
    borderStyle: "solid",
    borderRadius: 3,
    backgroundColor: "#fafafa",
  },
  cardTitle: {
    fontSize: 12,
    fontWeight: "bold",
    marginBottom: 8,
    color: "#333333",
    paddingBottom: 5,
    borderBottomWidth: 1,
    borderBottomColor: "#e0e0e0",
    borderBottomStyle: "solid",
  },
  cardContent: {
    paddingTop: 3,
  },
  cardText: {
    fontSize: 9,
    marginBottom: 3,
    lineHeight: 1.5,
  },
  alert: {
    marginVertical: 10,
    padding: 8,
    borderWidth: 1,
    borderStyle: "solid",
    borderRadius: 3,
  },
  alertText: {
    fontSize: 9,
    lineHeight: 1.5,
  },
  spacer: {
    height: 10,
  },
  // Chart styles
  chartDataContainer: {
    marginVertical: 8,
  },
  chartTotal: {
    flexDirection: "row",
    justifyContent: "space-between",
    borderBottomWidth: 1,
    borderBottomColor: "#e0e0e0",
    borderBottomStyle: "solid",
    paddingBottom: 6,
    marginBottom: 8,
  },
  chartTotalLabel: {
    fontSize: 12,
    fontWeight: "bold",
    color: "#333333",
  },
  chartTotalValue: {
    fontSize: 12,
    fontWeight: "bold",
    color: "#333333",
  },
  chartDataRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    paddingVertical: 4,
  },
  chartLabel: {
    fontSize: 9,
    color: "#666666",
    flex: 1,
  },
  chartValue: {
    fontSize: 9,
    color: "#666666",
    textAlign: "right",
  },
});

/**
 * Helper function to try to extract chart data from ApexCharts global instances
 * @param {HTMLElement} chartElement - DOM element containing the chart
 * @returns {Object|null} - Extracted chart data if available
 */
const tryExtractApexChartData = (chartElement) => {
  try {
    // This accesses ApexCharts' global instances which contain chart data
    // Note: This may not be accessible depending on the environment security
    if (window.ApexCharts && window.ApexCharts.getChartByID) {
      // Try to find the chart ID from the DOM
      const chartId = chartElement.id.split("-")[0];
      if (chartId) {
        const chart = window.ApexCharts.getChartByID(chartId);
        if (chart && chart.w) {
          // Extract data from the ApexCharts instance
          const series = chart.w.globals.series;
          const labels = chart.w.globals.labels;
          const chartType = chart.w.config.chart.type;

          return {
            type: chartType || "unknown",
            labels: labels || [],
            values: Array.isArray(series) ? series : [],
            totalValue: Array.isArray(series) ? series.reduce((a, b) => a + b, 0) : 0,
          };
        }
      }
    }
  } catch (err) {
    console.warn("Failed to extract ApexCharts data:", err);
  }
  return null;
};

/**
 * CippPDF component for rendering PDF content
 */
const CippPDF = ({ pageTitle = "CIPP Export", pageData = {}, captureScreenshot = null }) => {
  // Safety check for pageData structure
  const safePageData = {
    tables: Array.isArray(pageData.tables) ? pageData.tables : [],
    cards: Array.isArray(pageData.cards) ? pageData.cards : [],
    text: Array.isArray(pageData.text) ? pageData.text : [],
    pageInfo: pageData.pageInfo || {},
  };

  // Format date to be consistent
  const formattedDate = new Date().toLocaleString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

  return (
    <Document title={pageTitle} author="CIPP" creator="CIPP PDF Export" producer="CIPP">
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>{pageTitle}</Text>
          {safePageData.pageInfo && safePageData.pageInfo.tenant && (
            <Text style={styles.subtitle}>Tenant: {safePageData.pageInfo.tenant}</Text>
          )}
          <Text style={styles.timestamp}>Generated on: {formattedDate}</Text>
        </View>

        <View style={styles.content}>
          {captureScreenshot ? (
            <Image style={styles.image} src={captureScreenshot} />
          ) : (
            <>
              {/* Display cards first - they typically contain summary info */}
              {safePageData.cards && safePageData.cards.length > 0 && (
                <>
                  {safePageData.cards.map((card, cardIndex) => {
                    if (card.isBannerList) {
                      // Render Banner List Card with improved formatting
                      return (
                        <View style={styles.card} key={`banner-list-card-${cardIndex}`}>
                          {card.title && <Text style={styles.cardTitle}>{card.title}</Text>}
                          <View style={styles.cardContent}>
                            {card.items && card.items.length > 0 ? (
                              card.items.map((item, i) => (
                                <View
                                  key={`banner-item-${i}`}
                                  style={{
                                    marginBottom: 8,
                                    padding: 6,
                                    borderBottomWidth: i < card.items.length - 1 ? 1 : 0,
                                    borderBottomColor: "#e0e0e0",
                                    borderBottomStyle: "solid",
                                  }}
                                >
                                  {item.label && (
                                    <Text
                                      style={{
                                        fontWeight: "bold",
                                        fontSize: 10,
                                        color: "#1976d2",
                                        marginBottom: 2,
                                      }}
                                    >
                                      {item.label}
                                    </Text>
                                  )}
                                  {item.text && (
                                    <Text style={{ fontSize: 10, marginBottom: 1 }}>
                                      {item.text}
                                    </Text>
                                  )}
                                  {item.subtext && (
                                    <Text
                                      style={{ fontSize: 9, color: "#888", fontStyle: "italic" }}
                                    >
                                      {item.subtext}
                                    </Text>
                                  )}
                                </View>
                              ))
                            ) : (
                              <Text style={styles.cardText}>No items available.</Text>
                            )}
                          </View>
                        </View>
                      );
                    } else if (card.isInfoBar) {
                      // Render InfoBar with grid-like layout and clear separation
                      return (
                        <View style={styles.card} key={`infobar-card-${cardIndex}`}>
                          <View style={styles.cardContent}>
                            <View style={{ flexDirection: "row", flexWrap: "wrap", gap: 8 }}>
                              {card.items && card.items.length > 0 ? (
                                card.items.map((item, i) => (
                                  <View
                                    key={`infobar-item-${i}`}
                                    style={{
                                      minWidth: 90,
                                      marginRight: 18,
                                      marginBottom: 10,
                                      padding: 4,
                                      borderRadius: 2,
                                      backgroundColor: "#f0f4f8",
                                      flexDirection: "row",
                                      alignItems: "center",
                                    }}
                                  >
                                    {item.iconSrc && (
                                      <Image
                                        src={item.iconSrc}
                                        style={{ width: 14, height: 14, marginRight: 4 }}
                                      />
                                    )}
                                    {item.iconSvg && !item.iconSrc && (
                                      <Text style={{ marginRight: 4 }}>[icon]</Text>
                                    )}
                                    <View>
                                      {item.name && (
                                        <Text
                                          style={{ fontSize: 8, color: "#666", marginBottom: 1 }}
                                        >
                                          {item.name}
                                        </Text>
                                      )}
                                      {item.value ||
                                        (item.data && (
                                          <Text
                                            style={{
                                              fontSize: 13,
                                              fontWeight: "bold",
                                              color: "#333",
                                            }}
                                          >
                                            {item.value || item.data}
                                          </Text>
                                        ))}
                                    </View>
                                  </View>
                                ))
                              ) : (
                                <Text style={styles.cardText}>No info available.</Text>
                              )}
                            </View>
                          </View>
                        </View>
                      );
                    } else {
                      // Deduplicate card content before rendering
                      const uniqueContent = Array.isArray(card.content)
                        ? Array.from(new Set(card.content))
                        : [];
                      return (
                        <View style={styles.card} key={`card-${cardIndex}`}>
                          {card.title && <Text style={styles.cardTitle}>{card.title}</Text>}
                          <View style={styles.cardContent}>
                            {card.isChart && card.chartData ? (
                              <View>
                                {/* Chart data visualization in PDF */}
                                <View style={styles.chartDataContainer}>
                                  {/* Add total value if available */}
                                  {card.chartData.totalValue && (
                                    <View style={styles.chartTotal}>
                                      <Text style={styles.chartTotalLabel}>Total</Text>
                                      <Text style={styles.chartTotalValue}>
                                        {card.chartData.totalValue}
                                      </Text>
                                    </View>
                                  )}

                                  {/* Chart data items */}
                                  {card.chartData.labels.length > 0 ? (
                                    card.chartData.labels.map((label, i) => (
                                      <View key={`chart-data-${i}`} style={styles.chartDataRow}>
                                        <View
                                          style={{
                                            flexDirection: "row",
                                            alignItems: "center",
                                            flex: 1,
                                          }}
                                        >
                                          <View
                                            style={{
                                              width: 8,
                                              height: 8,
                                              borderRadius: 4,
                                              backgroundColor:
                                                i === 0
                                                  ? "#4caf50" // success - green
                                                  : i === 1
                                                  ? "#ff9800" // warning - orange
                                                  : i === 2
                                                  ? "#f44336" // error - red
                                                  : i === 3
                                                  ? "#2196f3" // info - blue
                                                  : i === 4
                                                  ? "#9c27b0" // purple
                                                  : i === 5
                                                  ? "#00bcd4" // cyan
                                                  : "#9e9e9e", // gray fallback
                                              marginRight: 6,
                                            }}
                                          />
                                          <Text style={styles.chartLabel}>{label}</Text>
                                        </View>
                                        <Text style={styles.chartValue}>
                                          {i < card.chartData.values.length
                                            ? card.chartData.values[i]
                                            : ""}
                                        </Text>
                                      </View>
                                    ))
                                  ) : (
                                    // Fallback for when we couldn't extract chart data
                                    <Text style={styles.cardText}>
                                      {card.chartData.type !== "unknown"
                                        ? `This is a ${card.chartData.type} chart. Please refer to the original page for full chart details.`
                                        : "Chart data visualization. Please refer to the original page for details."}
                                    </Text>
                                  )}
                                </View>
                              </View>
                            ) : (
                              // Regular card content, deduplicated
                              <>
                                {uniqueContent.map((item, itemIndex) => (
                                  <Text
                                    key={`card-item-${cardIndex}-${itemIndex}`}
                                    style={styles.cardText}
                                  >
                                    {item}
                                  </Text>
                                ))}
                              </>
                            )}
                          </View>
                        </View>
                      );
                    }
                  })}
                  <View style={styles.spacer} />
                </>
              )}

              {/* Display tables */}
              {safePageData.tables && safePageData.tables.length > 0 && (
                <>
                  {safePageData.tables.map((table, tableIndex) => (
                    <View style={styles.tableContainer} key={`table-${tableIndex}`}>
                      {table.title && <Text style={styles.tableTitle}>{table.title}</Text>}
                      <View style={styles.table}>
                        {/* Table header row */}
                        {table.headers && table.headers.length > 0 && (
                          <View style={styles.tableRowHeader}>
                            {table.headers.map((header, headerIndex) => (
                              <Text
                                key={`table-header-${tableIndex}-${headerIndex}`}
                                style={[
                                  styles.tableCellHeader,
                                  { flex: headerIndex === 0 ? 2 : 1 }, // Make first column wider if it's a label column
                                ]}
                              >
                                {header}
                              </Text>
                            ))}
                          </View>
                        )}

                        {/* Table rows */}
                        {table.rows &&
                          table.rows.map((row, rowIndex) => (
                            <View
                              key={`table-row-${tableIndex}-${rowIndex}`}
                              style={[
                                styles.tableRow,
                                rowIndex % 2 === 0 ? styles.tableRowEven : styles.tableRowOdd,
                              ]}
                            >
                              {row.map((cell, cellIndex) => {
                                // --- Render icon if present ---
                                if (cell && typeof cell === "object" && cell.icon) {
                                  return (
                                    <View
                                      key={`table-cell-${tableIndex}-${rowIndex}-${cellIndex}`}
                                      style={[
                                        styles.tableCell,
                                        {
                                          flex: cellIndex === 0 ? 2 : 1,
                                          flexDirection: "row",
                                          alignItems: "center",
                                        },
                                      ]}
                                    >
                                      {cell.icon.type === "img" ? (
                                        <Image
                                          src={cell.icon.src}
                                          style={{ width: 12, height: 12, marginRight: 4 }}
                                        />
                                      ) : (
                                        <View
                                          style={{
                                            width: 12,
                                            height: 12,
                                            backgroundColor: "#bbb",
                                            borderRadius: 2,
                                            marginRight: 4,
                                            alignItems: "center",
                                            justifyContent: "center",
                                          }}
                                        >
                                          <Text
                                            style={{
                                              fontSize: 7,
                                              color: "#fff",
                                              textAlign: "center",
                                            }}
                                          >
                                            icon
                                          </Text>
                                        </View>
                                      )}
                                      <Text style={{ fontSize: 8 }}>{cell.text || ""}</Text>
                                    </View>
                                  );
                                } else {
                                  return (
                                    <Text
                                      key={`table-cell-${tableIndex}-${rowIndex}-${cellIndex}`}
                                      style={[styles.tableCell, { flex: cellIndex === 0 ? 2 : 1 }]}
                                    >
                                      {cell}
                                    </Text>
                                  );
                                }
                              })}
                            </View>
                          ))}

                        {/* Show no data message if table is empty */}
                        {(!table.rows || table.rows.length === 0) && (
                          <View style={styles.tableRow}>
                            <Text style={[styles.tableCell, { textAlign: "center" }]}>
                              No data available
                            </Text>
                          </View>
                        )}
                      </View>
                      <View style={styles.spacer} />
                    </View>
                  ))}
                </>
              )}

              {/* Display text content (paragraphs, headings, etc.) */}
              {safePageData.text && safePageData.text.length > 0 && (
                <>
                  {safePageData.text.map((textItem, textIndex) => {
                    if (textItem.isAlert) {
                      // Render alerts with appropriate styling
                      const alertStyle = {
                        ...styles.alert,
                        borderColor:
                          textItem.type === "success"
                            ? "#4caf50"
                            : textItem.type === "error"
                            ? "#f44336"
                            : textItem.type === "warning"
                            ? "#ff9800"
                            : "#2196f3",
                        backgroundColor:
                          textItem.type === "success"
                            ? "#e8f5e9"
                            : textItem.type === "error"
                            ? "#ffebee"
                            : textItem.type === "warning"
                            ? "#fff3e0"
                            : "#e3f2fd",
                      };

                      return (
                        <View key={`text-${textIndex}`} style={alertStyle}>
                          <Text style={styles.alertText}>{textItem.text}</Text>
                        </View>
                      );
                    } else if (textItem.type === "heading") {
                      // Render headings
                      const fontSize = 18 - textItem.level * 2; // Decreasing font size for deeper heading levels
                      return (
                        <Text
                          key={`text-${textIndex}`}
                          style={[
                            styles.heading,
                            { fontSize, marginTop: textItem.level === 1 ? 16 : 12 },
                          ]}
                        >
                          {textItem.text}
                        </Text>
                      );
                    } else {
                      // Render regular paragraphs
                      return (
                        <Text key={`text-${textIndex}`} style={styles.paragraph}>
                          {textItem.text}
                        </Text>
                      );
                    }
                  })}
                </>
              )}

              {/* Show 'No Content' message if there's no data at all */}
              {(!safePageData.cards || safePageData.cards.length === 0) &&
                (!safePageData.tables || safePageData.tables.length === 0) &&
                (!safePageData.text || safePageData.text.length === 0) && (
                  <Text style={styles.noData}>No content available for export</Text>
                )}
            </>
          )}
        </View>

        <Text
          style={styles.footer}
          render={({ pageNumber, totalPages }) =>
            `CIPP Export | ${pageTitle} | Generated: ${formattedDate} | Page ${pageNumber} of ${totalPages}`
          }
          fixed
        />
      </Page>
    </Document>
  );
};

/**
 * Enhanced method to extract chart data from a chart element
 * @param {HTMLElement} card - The card element containing the chart
 * @param {HTMLElement} chartElement - The chart element
 * @param {string} chartType - The detected chart type
 * @returns {Object} - The extracted chart data
 */
const extractChartDataEnhanced = (card, chartElement, chartType) => {
  const chartData = {
    type: chartType,
    labels: [],
    values: [],
    totalValue: 0,
  };

  // Try to extract data from ApexCharts instance first (if available)
  try {
    if (window.ApexCharts && window.ApexCharts.getChartByID) {
      const chartId = chartElement.id.split("-")[0];
      if (chartId) {
        const chart = window.ApexCharts.getChartByID(chartId);
        if (chart && chart.w) {
          const series = chart.w.globals.series;
          const labels = chart.w.globals.labels;

          if (labels && labels.length > 0) {
            chartData.labels = [...labels];
            chartData.values = Array.isArray(series) ? series.map(String) : [];
            console.log("Successfully extracted chart data from ApexCharts instance");

            // If we successfully extracted data, return it
            return chartData;
          }
        }
      }
    }
  } catch (err) {
    console.warn("Failed to extract chart data from ApexCharts:", err);
  }

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

/**
 * Modified extractPageData function with enhanced chart detection
 */
const extractPageData = async (element) => {
  if (!element) return { tables: [], text: [], cards: [], pageInfo: {} };

  const data = {
    tables: [],
    text: [],
    cards: [],
    pageInfo: {
      title: document.title || "CIPP",
      url: window.location.href,
      timestamp: new Date().toLocaleString(),
      tenant: "",
    },
  };

  try {
    // Try to extract page-level information first
    const pageTitle = element.querySelector("h1, h2, .MuiTypography-h4");
    if (pageTitle) {
      data.pageInfo.pageTitle = pageTitle.textContent.trim();
    }

    // Try to find tenant information
    const tenantInfo = document.querySelector("[aria-label='Select Tenant']");
    if (tenantInfo) {
      data.pageInfo.tenant = tenantInfo.textContent.trim();
    }

    // Extract Material UI tables (MRT tables used in CippDataTable)
    const mrtTables = element.querySelectorAll(".MuiTable-root");
    for (const table of mrtTables) {
      const tableData = { headers: [], rows: [], title: "" };
      // Look for table title in parent components
      let parentEl = table.closest(".MuiCard-root");
      if (parentEl) {
        const titleEl = parentEl.querySelector(".MuiCardHeader-title");
        if (titleEl) {
          tableData.title = titleEl.textContent.trim();
        }
      }
      // Extract headers from MUI table, always include unless 'Actions' or 'RowKey'
      const headerCells = table.querySelectorAll(".MuiTableHead-root th");
      headerCells.forEach((cell) => {
        const headerText = cell.textContent.trim();
        if (
          headerText &&
          !headerText.toLowerCase().includes("action") &&
          !headerText.includes("RowKey")
        ) {
          tableData.headers.push(headerText);
        }
      });
      if (tableData.headers.length === 0) continue;
      // Extract rows from MUI table, skipping cells with buttons/inputs but keeping alignment
      const rows = table.querySelectorAll(".MuiTableBody-root tr");
      for (const row of rows) {
        const rowData = [];
        const cells = row.querySelectorAll("td");
        for (let colIdx = 0; colIdx < headerCells.length; colIdx++) {
          const cell = headerCells[colIdx];
          const headerText = cell.textContent.trim();
          if (headerText.toLowerCase().includes("action") || headerText.includes("RowKey")) {
            // Always push empty string to maintain alignment
            rowData.push("");
            continue;
          }
          const dataCell = cells[colIdx];
          if (!dataCell) {
            rowData.push("");
            continue;
          }
          // Skip checkboxes, action buttons, and interactive elements, but always push empty string to maintain alignment
          if (
            dataCell.querySelector(
              'button, input[type="checkbox"], input, textarea, [role="button"], [role="search"]'
            )
          ) {
            rowData.push("");
            continue;
          }
          // --- Enhanced chip list extraction (await SVG PNG conversion) ---
          const chipNodes = dataCell.querySelectorAll(".MuiChip-root");
          if (chipNodes.length > 0) {
            const chips = [];
            for (const chip of chipNodes) {
              const labelSpan = chip.querySelector(".MuiChip-label");
              let label = labelSpan ? labelSpan.textContent.trim() : chip.textContent.trim();
              let icon = null;
              const img = chip.querySelector("img");
              if (img && img.src) {
                icon = { type: "img", src: img.src };
              } else {
                const svg = chip.querySelector("svg");
                if (svg) {
                  // Await PNG conversion for SVGs
                  const pngUrl = await svgToPngDataUrl(svg);
                  icon = pngUrl ? { type: "img", src: pngUrl } : { type: "svg" };
                }
              }
              chips.push(icon ? { label, icon } : label);
            }
            // --- Fix: Always push a single value for the cell ---
            if (chips.every((c) => typeof c === "string")) {
              rowData.push(chips.join(", "));
            } else {
              // At least one chip has an icon: use the first icon and join all labels
              const firstWithIcon = chips.find((c) => typeof c === "object" && c.icon);
              const allLabels = chips.map((c) => (typeof c === "string" ? c : c.label)).join(", ");
              if (firstWithIcon) {
                rowData.push({ icon: firstWithIcon.icon, text: allLabels });
              } else {
                rowData.push(allLabels); // fallback, should not happen
              }
            }
            continue;
          }
          // --- Fallback: If no .MuiChip-root, but .MuiChip-label(s) exist, extract all labels ---
          const chipLabels = dataCell.querySelectorAll(".MuiChip-label");
          if (chipLabels.length > 0) {
            const allLabels = Array.from(chipLabels)
              .map((lbl) => lbl.textContent.trim())
              .filter(Boolean);
            if (allLabels.length > 0) {
              rowData.push(allLabels.join(", "));
              continue;
            }
          }
          // If the cell contains a phone/email chip, extract its text
          const chipLabel = dataCell.querySelector(".MuiChip-label");
          if (chipLabel) {
            rowData.push(chipLabel.textContent.trim());
            continue;
          }
          // --- End phone/email chip extraction ---
          // Icon extraction for regular cells
          let icon = null;
          const img = dataCell.querySelector("img");
          if (img && img.src) {
            icon = { type: "img", src: img.src };
          } else {
            const svg = dataCell.querySelector("svg");
            if (svg) {
              const pngUrl = await svgToPngDataUrl(svg);
              icon = pngUrl ? { type: "img", src: pngUrl } : { type: "svg" };
            }
          }
          let cellContent = dataCell.textContent.trim();
          if (icon && cellContent) {
            rowData.push({ icon, text: cellContent });
          } else if (icon) {
            rowData.push({ icon });
          } else {
            rowData.push(cellContent);
          }
        }
        // Always push the row, even if empty, to maintain alignment
        tableData.rows.push(rowData);
      }
      if (tableData.rows.length > 0) {
        data.tables.push(tableData);
      }
    }

    // Extract card content - MUI cards are often used for summary info
    const cards = element.querySelectorAll(".MuiCard-root");
    cards.forEach((card) => {
      // Avoid cards that contain tables we've already processed
      if (card.querySelector("table") === null) {
        const cardTitle = card.querySelector(".MuiCardHeader-title");
        const cardContent = card.querySelector(".MuiCardContent-root");

        if (cardContent) {
          // --- InfoBar-like card detection ---
          const gridItems = cardContent.querySelectorAll(".MuiGrid-root, .MuiStack-root");
          let infoBarItems = [];
          gridItems.forEach((item) => {
            // Exclude items with interactive elements
            if (item.querySelector('button, input, textarea, [role="button"], [role="search"]'))
              return;
            // Find .MuiTypography-overline and .MuiTypography-h6 anywhere inside this item (deep search)
            const nameEl = item.querySelector(".MuiTypography-overline");
            const dataEl = item.querySelector(".MuiTypography-h6");
            if (nameEl && dataEl) {
              let icon = null;
              const img = item.querySelector("img");
              if (img && img.src) {
                icon = { type: "img", src: img.src };
              } else {
                const svg = item.querySelector("svg");
                if (svg) {
                  // Await PNG conversion for SVGs (if needed)
                }
              }
              infoBarItems.push({
                name: nameEl.textContent.trim(),
                data: dataEl.textContent.trim(),
                icon,
              });
            }
          });
          if (infoBarItems.length > 1) {
            data.cards.push({
              isInfoBar: true,
              items: infoBarItems,
            });
            return; // skip regular card extraction for this card
          }

          const cardData = {
            title: cardTitle ? cardTitle.textContent.trim() : "",
            content: [],
            isChart: false,
            chartData: null,
          };

          // Check if this card contains a chart (ApexChart used by CippChartCard)
          const chartElement = card.querySelector(".apexcharts-canvas");
          if (chartElement) {
            // This is a chart card, handle it specially
            cardData.isChart = true;

            // Get chart type from data attributes or class names
            let chartType = "unknown";
            if (
              chartElement.classList.contains("apexcharts-pie") ||
              card.querySelector(".apexcharts-pie")
            ) {
              chartType = "donut"; // or pie
            } else if (
              chartElement.classList.contains("apexcharts-bar") ||
              card.querySelector(".apexcharts-bar")
            ) {
              chartType = "bar";
            } else if (
              chartElement.classList.contains("apexcharts-line") ||
              card.querySelector(".apexcharts-line")
            ) {
              chartType = "line";
            }

            // Try to get chart type from title as a fallback
            if (chartType === "unknown" && cardTitle) {
              const title = cardTitle.textContent.toLowerCase();
              if (title.includes("pie") || title.includes("donut")) {
                chartType = "donut";
              } else if (title.includes("bar")) {
                chartType = "bar";
              } else if (title.includes("line")) {
                chartType = "line";
              }
            }

            // Extract chart data using our enhanced method
            cardData.chartData = extractChartDataEnhanced(card, chartElement, chartType);

            // Also add all the text content for backwards compatibility
            const textItems = cardContent.querySelectorAll("p, h1, h2, h3, h4, h5, h6, div, span");
            textItems.forEach((item) => {
              const text = item.textContent.trim();
              if (text && text.length > 0) {
                cardData.content.push(text);
              }
            });
          } else {
            // Regular, non-chart card - just extract text content
            const textItems = cardContent.querySelectorAll("p, h1, h2, h3, h4, h5, h6, div, span");
            textItems.forEach((item) => {
              const text = item.textContent.trim();
              if (text && text.length > 0) {
                cardData.content.push(text);
              }
            });
          }

          data.cards.push(cardData);
        }
      }
    });

    // In extractPageData, add support for CippBannerListCard and CippInfoBar
    // Banner List Card extraction
    const bannerListCards = element.querySelectorAll(
      ".CippBannerListCard-root, [data-cipp-banner-list-card]"
    );
    bannerListCards.forEach((card) => {
      const cardTitle = card.querySelector(".MuiCardHeader-title, h2, h3, h4");
      const cardData = {
        title: cardTitle ? cardTitle.textContent.trim() : "",
        content: [],
        isBannerList: true,
        items: [],
      };
      // Extract banner items
      const items = card.querySelectorAll("li");
      items.forEach((li) => {
        const label = li.querySelector("h5, .MuiTypography-h5");
        const text = li.querySelector("h6, .MuiTypography-h6");
        const subtext = li.querySelector(".MuiTypography-body2, .MuiTypography-caption");
        cardData.items.push({
          label: label ? label.textContent.trim() : "",
          text: text ? text.textContent.trim() : "",
          subtext: subtext ? subtext.textContent.trim() : "",
        });
      });
      if (cardData.items.length > 0) {
        data.cards.push(cardData);
      }
    });
    // InfoBar extraction
    const infoBars = element.querySelectorAll(".CippInfoBar-root, [data-cipp-info-bar]");
    for (const bar of infoBars) {
      const infoData = {
        isInfoBar: true,
        items: [],
      };
      let items = bar.querySelectorAll('.MuiGrid-root, [role="gridcell"]');
      if (!items.length) {
        items = bar.children;
      }
      for (const item of items) {
        if (item.querySelector('button, input, textarea, [role="button"], [role="search"]')) {
          continue;
        }
        let name = item.querySelector(".MuiTypography-overline");
        let value = item.querySelector(".MuiTypography-h6");
        let icon = null;
        const img = item.querySelector("img");
        if (img && img.src) {
          icon = { type: "img", src: img.src };
        } else {
          const svg = item.querySelector("svg");
          if (svg) {
            const pngUrl = await svgToPngDataUrl(svg);
            if (pngUrl) {
              icon = { type: "img", src: pngUrl };
            } else {
              icon = { type: "svg" };
            }
          }
        }
        // Always set both value and data for compatibility
        const valueText = value ? value.textContent.trim() : item.textContent.trim();
        if ((!name || !value) && item.textContent.trim()) {
          infoData.items.push({
            name: name ? name.textContent.trim() : "",
            value: valueText,
            data: valueText,
            icon,
          });
        } else if (name || value) {
          infoData.items.push({
            name: name ? name.textContent.trim() : "",
            value: value ? value.textContent.trim() : "",
            data: value ? value.textContent.trim() : "",
            icon,
          });
        }
      }
      if (infoData.items.length > 0) {
        data.cards.push(infoData);
      }
    }

    // Extract any standalone headings and text paragraphs
    const headings = element.querySelectorAll("h1, h2, h3, h4, h5, h6");
    headings.forEach((heading) => {
      // Ensure this heading isn't inside a card or table we've already processed
      if (!heading.closest(".MuiCard-root") && !heading.closest("table")) {
        data.text.push({
          type: "heading",
          level: parseInt(heading.tagName.substring(1)),
          text: heading.textContent.trim(),
        });
      }
    });

    const paragraphs = element.querySelectorAll("p");
    paragraphs.forEach((para) => {
      // Ensure this paragraph isn't inside a card or table we've already processed
      if (!para.closest(".MuiCard-root") && !para.closest("table")) {
        data.text.push({
          type: "paragraph",
          text: para.textContent.trim(),
        });
      }
    });

    return data;
  } catch (error) {
    console.error("Error extracting page data:", error);
    return { tables: [], text: [], cards: [], pageInfo: {} };
  }
};

// Async helper to convert SVG to PNG data URL using html2canvas, with caching
const svgPngCache = new Map();
async function svgToPngDataUrl(svgElement) {
  if (!svgElement) return null;
  const svgKey = svgElement.outerHTML;
  if (svgPngCache.has(svgKey)) {
    return svgPngCache.get(svgKey);
  }
  try {
    const html2canvas = (await import("html2canvas")).default;
    const clone = svgElement.cloneNode(true);
    // --- Force light mode styling on SVG ---
    // Remove dark mode classes if present
    clone.classList.remove("MuiSvgIcon-root-dark", "dark-mode", "theme-dark");
    // Set background to white and text/fill/stroke to dark for all children
    clone.style.background = "#fff";
    // Recursively set fill and stroke for all elements in the SVG
    function setLightModeStyles(el) {
      if (el.tagName && el.tagName.toLowerCase() !== "svg") {
        if (el.hasAttribute("fill")) el.setAttribute("fill", "#222");
        if (el.hasAttribute("stroke")) el.setAttribute("stroke", "#222");
        if (el.style) {
          el.style.fill = "#222";
          el.style.stroke = "#222";
        }
      }
      for (const child of el.children || []) {
        setLightModeStyles(child);
      }
    }
    setLightModeStyles(clone);
    const wrapper = document.createElement("div");
    wrapper.appendChild(clone);
    document.body.appendChild(wrapper);
    const canvas = await html2canvas(wrapper, {
      backgroundColor: "#fff",
      useCORS: true,
      logging: false,
      allowTaint: true,
      width: svgElement.clientWidth || 24,
      height: svgElement.clientHeight || 24,
      scale: 1,
    });
    document.body.removeChild(wrapper);
    const dataUrl = canvas.toDataURL("image/png");
    svgPngCache.set(svgKey, dataUrl);
    return dataUrl;
  } catch (e) {
    console.warn("Failed to convert SVG to PNG:", e);
    return null;
  }
}

/**
 * PDF Export Dialog Component
 */
const PdfExportDialog = ({ open, onClose, pageTitle, pageData, captureScreenshot }) => {
  const [saving, setSaving] = useState(false);
  const [viewMode, setViewMode] = useState("preview");

  const handleSave = useCallback(async () => {
    try {
      setSaving(true);
      const blob = await pdf(
        <CippPDF pageTitle={pageTitle} pageData={pageData} captureScreenshot={captureScreenshot} />
      ).toBlob();
      saveAs(
        blob,
        `CIPP-Export-${pageTitle.replace(/[^a-z0-9]/gi, "-").toLowerCase()}-${
          new Date().toISOString().split("T")[0]
        }.pdf`
      );
      toast.success("PDF saved successfully");
      setTimeout(() => onClose(), 500);
    } catch (err) {
      console.error("Error saving PDF:", err);
      toast.error("Failed to save PDF");
    } finally {
      setSaving(false);
    }
  }, [pageTitle, pageData, captureScreenshot, onClose]);

  return (
    <Dialog
      open={open}
      onClose={onClose}
      fullWidth
      maxWidth="lg"
      sx={{
        "& .MuiDialog-paper": {
          height: "90vh",
        },
      }}
    >
      <DialogTitle
        sx={{ m: 0, p: 2, display: "flex", alignItems: "center", justifyContent: "space-between" }}
      >
        <div>Export {pageTitle} as PDF</div>
        <div>
          <button
            onClick={() => setViewMode("preview")}
            style={{
              padding: "6px 12px",
              marginRight: 8,
              backgroundColor: viewMode === "preview" ? "#1976d2" : "#f5f5f5",
              color: viewMode === "preview" ? "#fff" : "#333",
              border: "none",
              borderRadius: 4,
              cursor: "pointer",
            }}
          >
            Preview
          </button>
          <button
            onClick={handleSave}
            disabled={saving}
            style={{
              padding: "6px 12px",
              marginRight: 8,
              backgroundColor: saving ? "#cccccc" : "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: 4,
              cursor: saving ? "default" : "pointer",
            }}
          >
            {saving ? "Saving..." : "Save PDF"}
          </button>
          <IconButton
            aria-label="close"
            onClick={onClose}
            sx={{
              color: (theme) => theme.palette.grey[500],
            }}
          >
            <CloseIcon />
          </IconButton>
        </div>
      </DialogTitle>
      <DialogContent dividers sx={{ p: 0, display: "flex", flexDirection: "column" }}>
        <div
          style={{
            flex: 1,
            overflow: "auto",
            display: "flex",
            justifyContent: "center",
            backgroundColor: "#f5f5f5",
          }}
        >
          <PDFViewer width="100%" height="100%" style={{ border: "none" }}>
            <CippPDF
              pageTitle={pageTitle}
              pageData={pageData}
              captureScreenshot={captureScreenshot}
            />
          </PDFViewer>
        </div>
      </DialogContent>
    </Dialog>
  );
};

/**
 * Main CippPrintPage component for exporting PDF
 */
const CippPrintPage = {
  id: "export-pdf",
  name: "Export as PDF",
  handler: async () => {
    try {
      // Show initial toast notification
      const loadingToast = toast.loading("Preparing PDF export...");

      // Get the current page title
      const pageTitle = document.title || "CIPP";
      const pageName = pageTitle.replace("CIPP - ", "").trim();

      // Find the main content element with appropriate selectors
      const contentSelectors = [
        ".LayoutContainer > div",
        ".MuiContainer-root",
        "main",
        ".container",
        "[class^='MuiBox-root']",
      ];

      let contentElement = null;
      for (const selector of contentSelectors) {
        const element = document.querySelector(selector);
        if (element) {
          console.log(`Found content using selector: ${selector}`);
          contentElement = element;
          break;
        }
      }

      if (!contentElement) {
        console.warn("Could not find specific content container, using body");
        contentElement = document.body;
      }

      // Extract structured data from the page
      const pageData = await extractPageData(contentElement);

      // Only attempt to capture screenshot if structured data extraction was limited
      const hasLimitedData =
        pageData.tables.length === 0 &&
        pageData.cards.filter((c) => !c.isChart).length === 0 &&
        pageData.text.length === 0;

      let screenshot = null;
      if (hasLimitedData) {
        try {
          // Check if html2canvas is available
          const html2canvas = await import("html2canvas").catch(() => null);
          if (html2canvas && html2canvas.default) {
            // Capture screenshot with reasonable quality settings
            const canvas = await html2canvas.default(contentElement, {
              scale: 1,
              useCORS: true,
              logging: false,
              allowTaint: true,
              backgroundColor: "#ffffff",
              ignoreElements: (element) => {
                // Exclude certain elements from screenshot
                return (
                  element.classList.contains("MuiSpeedDial-root") ||
                  element.id === "cipp-pdf-dialog-container"
                );
              },
            });
            screenshot = canvas.toDataURL("image/png");
            console.log("Screenshot captured successfully");
          }
        } catch (screenshotError) {
          console.warn("Failed to capture page screenshot:", screenshotError);
        }
      }

      // Close loading toast
      toast.dismiss(loadingToast);

      // Create container for dialog
      const dialogContainer = document.createElement("div");
      dialogContainer.id = "cipp-pdf-dialog-container";
      document.body.appendChild(dialogContainer);

      // Create React root and render dialog
      const ReactDOM = await import("react-dom/client");
      const root = ReactDOM.createRoot(dialogContainer);

      // Set up close handler
      const handleClose = () => {
        root.unmount();
        document.body.removeChild(dialogContainer);
      };

      // Render PDF preview dialog
      root.render(
        <PdfExportDialog
          open={true}
          onClose={handleClose}
          pageTitle={pageName}
          pageData={pageData}
          captureScreenshot={screenshot}
        />
      );
    } catch (error) {
      console.error("Error in PDF export:", error);
      toast.error("Failed to export page as PDF");
    }
  },
};

export default CippPrintPage;
