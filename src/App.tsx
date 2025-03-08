/* eslint-disable @typescript-eslint/no-explicit-any */
import { useState, useEffect } from "react";
import {
  Box,
  Button,
  Container,
  TextField,
  Typography,
  Paper,
  Alert,
  CircularProgress,
  IconButton,
  Collapse,
  Divider,
  Stack,
  Tooltip,
  Switch,
  FormControlLabel,
} from "@mui/material";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import UploadFileIcon from "@mui/icons-material/UploadFile";
import SettingsIcon from "@mui/icons-material/Settings";
import EmailIcon from "@mui/icons-material/Email";
import TranslateIcon from "@mui/icons-material/Translate";
import * as XLSX from "xlsx-js-style";
import "./App.css";

// Translations
const translations = {
  en: {
    title: "Dorm Score Formatter",
    subtitle: "Convert and format your dorm scoring data with ease",
    step1: "Step 1:",
    step2: "Step 2:",
    step3: "Step 3:",
    optional: "Optional:",
    fileSelection: "File Selection",
    selectFiles: "Select CSV Files",
    filesSelected: "file(s) selected",
    contactInfo: "Contact Information",
    emailPrefix: "THU Email Prefix",
    emailHelperText:
      "Enter your THU email prefix (without @mails.tsinghua.edu.cn)",
    emailPlaceholder: "e.g., yunyang-21",
    advancedOptions: "Advanced Options",
    rowsPerPage: "Rows Per Page",
    rowsHelperText: "Number of rows to display per page in the Excel file",
    processFiles: "Process Files",
    processing: "Processing...",
    darkMode: "Dark Mode",
    language: "中文",
  },
  zh: {
    title: "宿舍评分格式化工具",
    subtitle: "轻松转换和格式化宿舍评分数据",
    step1: "步骤 1 :",
    step2: "步骤 2 :",
    step3: "步骤 3 :",
    optional: "可选 :",
    fileSelection: "选择文件",
    selectFiles: "选择CSV文件",
    filesSelected: "个文件已选择",
    contactInfo: "联系方式",
    emailPrefix: "清华邮箱前缀",
    emailHelperText: "输入您的清华邮箱前缀（不含 @mails.tsinghua.edu.cn）",
    emailPlaceholder: "例如：yunyang-21",
    advancedOptions: "高级选项",
    rowsPerPage: "每页行数",
    rowsHelperText: "Excel文件中每页包含的行数",
    processFiles: "处理文件",
    processing: "处理中...",
    darkMode: "深色模式",
    language: "En",
  },
};

interface AppProps {
  darkMode: boolean;
  onDarkModeChange: (darkMode: boolean) => void;
}

function App({ darkMode, onDarkModeChange }: AppProps) {
  const [files, setFiles] = useState<FileList | null>(null);
  const [emailPrefix, setEmailPrefix] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [rowsPerPage, setRowsPerPage] = useState(50);
  const [language, setLanguage] = useState<"en" | "zh">(() => {
    const saved = localStorage.getItem("language");
    return saved === "en" || saved === "zh" ? saved : "en";
  });

  // Save language preference when it changes
  useEffect(() => {
    localStorage.setItem("language", language);
  }, [language]);

  const t = translations[language];

  const processFiles = async () => {
    console.log("Starting file processing...");
    console.log("Current state:", {
      filesCount: files?.length || 0,
      emailPrefix,
    });

    if (!files || files.length === 0) {
      console.warn("No files selected");
      setError("Please select CSV files to process");
      return;
    }

    if (!emailPrefix) {
      console.warn("No email prefix provided");
      setError("Please enter your THU email prefix");
      return;
    }

    setLoading(true);
    setError(null);
    setSuccess(null);

    try {
      const fileData: { file: string; data: any[] }[] = [];
      let buildingNumber = null;
      const weekCounts = new Map<number, number>();

      // First pass: load all files and count weeks
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        console.log(`Processing file ${i + 1}/${files.length}:`, file.name);

        if (
          !file.name.startsWith("WeekScoreManage_") ||
          !file.name.endsWith(".csv")
        ) {
          console.warn(`Skipping invalid file: ${file.name}`);
          continue;
        }

        const data = await readCSVFile(file);
        console.log(`File ${file.name} data:`, {
          rowCount: data.length,
          firstRow: data[0],
          lastRow: data[data.length - 1],
        });

        if (!data || data.length === 0) {
          console.warn(`No data found in file: ${file.name}`);
          continue;
        }

        // Set building number from first valid file
        if (buildingNumber === null) {
          buildingNumber = data[0]["楼号"];
        } else if (buildingNumber !== data[0]["楼号"]) {
          throw new Error(
            `Building number mismatch in file ${file.name}. Expected ${buildingNumber}, found ${data[0]["楼号"]}`
          );
        }

        // Count occurrences of each week
        data.forEach((row) => {
          const weekStr = row["周"].toString();
          const weekMatch = weekStr.match(/第(\d+)周/);
          if (weekMatch) {
            const week = parseInt(weekMatch[1]);
            weekCounts.set(week, (weekCounts.get(week) || 0) + 1);
          } else {
            console.warn(`Invalid week format in row:`, { weekStr, row });
          }
        });

        fileData.push({ file: file.name, data });
      }

      if (fileData.length === 0) {
        throw new Error("No valid files found");
      }

      if (weekCounts.size === 0) {
        throw new Error("No valid week numbers found in any files");
      }

      // Find the latest week
      const targetWeek = Math.max(...Array.from(weekCounts.keys()));

      console.log("Week analysis:", {
        availableWeeks: Array.from(weekCounts.keys()).sort((a, b) => a - b),
        latestWeek: targetWeek,
        dataPointsInLatestWeek: weekCounts.get(targetWeek),
        allWeekCounts: Object.fromEntries(weekCounts),
      });

      // Check which files are missing data for the latest week
      const filesWithMissingData = fileData
        .filter(
          ({ data }) =>
            !data.some((row) => {
              const weekMatch = row["周"].toString().match(/第(\d+)周/);
              return weekMatch && parseInt(weekMatch[1]) === targetWeek;
            })
        )
        .map(({ file }) => file);

      if (filesWithMissingData.length > 0) {
        console.warn(
          `Files missing data for week ${targetWeek} (latest week):`,
          filesWithMissingData
        );
        setError(
          `Warning: The following files don't contain data for week ${targetWeek} (latest week): ${filesWithMissingData.join(
            ", "
          )}`
        );
      }

      // Filter data for the latest week and combine
      const allData = fileData.flatMap(({ data }) =>
        data.filter((row) => {
          const weekMatch = row["周"].toString().match(/第(\d+)周/);
          return weekMatch && parseInt(weekMatch[1]) === targetWeek;
        })
      );

      console.log(
        `Filtered data for latest week (week ${targetWeek}), total rows:`,
        allData.length
      );

      if (allData.length === 0) {
        throw new Error(`No data found for week ${targetWeek} (latest week)`);
      }

      // Process the combined data
      console.log("Starting data processing...");
      const processedData = processData(allData);
      console.log(
        "Data processing complete. Processed rows:",
        processedData.length
      );

      // Create Excel file
      console.log("Creating Excel workbook...");
      const workbook = createExcelWorkbook(processedData, emailPrefix);
      console.log("Excel workbook created");

      // Generate the Excel file
      const excelFileName = `${buildingNumber}${targetWeek}.xlsx`;
      console.log("Saving Excel file:", excelFileName);
      XLSX.writeFile(workbook, excelFileName);
      console.log("Excel file saved successfully");

      setSuccess(
        filesWithMissingData.length > 0
          ? `Files processed successfully for week ${targetWeek}, but some files were missing data.`
          : "Files processed successfully!"
      );
    } catch (err) {
      console.error("Error during processing:", err);
      setError(
        err instanceof Error
          ? err.message
          : "An error occurred while processing the files"
      );
    } finally {
      setLoading(false);
      console.log("Processing complete");
    }
  };

  const readCSVFile = (file: File): Promise<any[]> => {
    console.log(`Reading CSV file: ${file.name}`);
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          console.log(`File ${file.name} loaded, decoding with GB2312...`);
          const arrayBuffer = e.target?.result as ArrayBuffer;

          // Create a Uint8Array from the array buffer
          const uint8Array = new Uint8Array(arrayBuffer);

          // Try to decode with GB2312 (using GBK as it's a superset of GB2312)
          let decoder;
          try {
            decoder = new TextDecoder("gbk");
          } catch {
            console.warn("GBK decoder not available, falling back to GB18030");
            decoder = new TextDecoder("GB18030");
          }

          const decodedText = decoder.decode(uint8Array);
          console.log(`File ${file.name} decoded successfully`);

          const workbook = XLSX.read(decodedText, { type: "string" });
          console.log(`File ${file.name} parsed, sheets:`, workbook.SheetNames);

          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);
          console.log(
            `File ${file.name} converted to JSON, rows:`,
            jsonData.length
          );

          // Log first row to verify encoding
          if (jsonData.length > 0) {
            console.log("First row sample:", jsonData[0]);
          }

          resolve(jsonData);
        } catch (err) {
          console.error(`Error reading file ${file.name}:`, err);
          reject(err);
        }
      };
      reader.onerror = (err) => {
        console.error(`File read error for ${file.name}:`, err);
        reject(err);
      };
      reader.readAsArrayBuffer(file);
    });
  };

  const processData = (data: any[]) => {
    console.log("Processing data, initial count:", data.length);
    // Remove duplicates and sort
    const uniqueData = data.reduce((acc: any[], curr: any) => {
      const exists = acc.find(
        (item) =>
          item["楼号"] === curr["楼号"] &&
          item["房间"] === curr["房间"] &&
          item["床位"] === curr["床位"]
      );
      if (!exists) {
        acc.push(curr);
      } else {
        console.log("Duplicate entry found:", {
          building: curr["楼号"],
          room: curr["房间"],
          bed: curr["床位"],
        });
      }
      return acc;
    }, []);

    console.log("Duplicates removed, count:", uniqueData.length);

    const sortedData = uniqueData.sort((a: any, b: any) => {
      if (a["房间"] === b["房间"]) {
        return a["床位"] - b["床位"];
      }
      return a["房间"].localeCompare(b["房间"]);
    });

    console.log("Data sorted, final count:", sortedData.length);
    return sortedData;
  };

  const createExcelWorkbook = (data: any[], emailPrefix: string) => {
    console.log("Creating Excel workbook with data rows:", data.length);
    const wb = XLSX.utils.book_new();

    // Initialize worksheet with empty array
    const ws = XLSX.utils.aoa_to_sheet([[""]]); // Initialize with at least one cell

    // Set column widths for all columns (A-H)
    ws["!cols"] = [
      { width: 6 }, // A
      { width: 6 }, // B
      { width: 6 }, // C
      { width: 25 }, // D
      { width: 6 }, // E
      { width: 6 }, // F
      { width: 6 }, // G
      { width: 25 }, // H
    ];

    // Add title and headers with merging
    const title = `${data[0]["楼号"]}${data[0]["周"]}`;
    console.log("Adding title:", title);

    // Merge cells for title and headers (A1:H1, A2:H2, A3:H3)
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, // A1:H1
      { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }, // A2:H2
      { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } }, // A3:H3
    ];

    // Helper function to set cell value and style
    const addFormattedCell = (cellRef: string, _row: number, value: string) => {
      // Set the cell value
      ws[cellRef] = {
        v: value,
        t: "s",
        s: {
          alignment: {
            vertical: "center",
            horizontal: "center",
            wrapText: true,
          },
          font: {
            name: "宋体",
            sz: 11,
          },
          border: {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" },
          },
        },
      };
    };

    // Add title rows
    addFormattedCell("A1", 1, title);
    addFormattedCell(
      "A2",
      2,
      "勤工大队楼层长分队统一意见邮箱：thu.lczh@gmail.com"
    );
    addFormattedCell(
      "A3",
      3,
      `如有疑问请联系学生楼长：${emailPrefix}@mails.tsinghua.edu.cn 或登陆家园网查询具体成绩`
    );

    // Add headers
    const headers = ["房间", "床位", "总分", "整改意见"];
    headers.forEach((header, idx) => {
      addFormattedCell(XLSX.utils.encode_cell({ r: 3, c: idx }), 4, header);
      addFormattedCell(XLSX.utils.encode_cell({ r: 3, c: idx + 4 }), 4, header);
    });

    // Process data rows
    console.log("Processing data for Excel...");
    const ROW_PER_PAGE = rowsPerPage;
    let currentRow = 0;
    let currentCol = 0;
    let rowsOnPage = 4;
    let page = 0;

    // Calculate max row and column for reference range
    let maxRow = 4;
    let maxCol = 7; // H column (0-based index)

    data.forEach((row) => {
      if (rowsOnPage === ROW_PER_PAGE) {
        if (currentCol === 0) {
          currentCol = 4;
          rowsOnPage = page === 0 ? 4 : 0;
        } else {
          page++;
          currentCol = 0;
          rowsOnPage = 0;
        }
      }

      currentRow = page * ROW_PER_PAGE + rowsOnPage;
      const rowData = [row["房间"], row["床位"], row["总分"], row["整改意见"]];
      rowData.forEach((value, idx) => {
        const r = currentRow;
        const c = currentCol + idx;
        const cell = XLSX.utils.encode_cell({ r, c });

        // Update max row/col for reference range
        maxRow = Math.max(maxRow, r);
        maxCol = Math.max(maxCol, c);

        const fontSize =
          idx === 3 && value?.toString().length > 10
            ? 11 - Math.max(0, (value.toString().length - 10) / 2)
            : 11;

        // Set cell with value and style
        ws[cell] = {
          v: value || "",
          t: "s",
          s: {
            alignment: {
              vertical: "center",
              horizontal: "center",
              wrapText: true,
            },
            font: {
              name: "宋体",
              sz: fontSize,
            },
            border: {
              top: { style: "thin" },
              bottom: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" },
            },
          },
        };
      });

      rowsOnPage++;
    });

    // Set the worksheet reference range
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxCol },
    });

    console.log("Final worksheet state:", {
      ref: ws["!ref"],
      merges: ws["!merges"],
      cols: ws["!cols"],
    });

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    return wb;
  };

  return (
    <Container maxWidth="md" sx={{ py: 6 }}>
      <Stack spacing={4}>
        {/* Header Section with Theme and Language Controls */}
        <Box
          sx={{
            display: "flex",
            justifyContent: "flex-end",
            gap: 2,
            alignItems: "center",
          }}
        >
          <FormControlLabel
            control={
              <Switch
                checked={darkMode}
                onChange={(e) => onDarkModeChange(e.target.checked)}
              />
            }
            label={t.darkMode}
          />
          <Button
            startIcon={<TranslateIcon />}
            onClick={() => setLanguage(language === "en" ? "zh" : "en")}
          >
            {t.language}
          </Button>
        </Box>

        {/* Title Section */}
        <Box sx={{ textAlign: "center", mb: 2 }}>
          <Typography
            variant="h3"
            component="h1"
            gutterBottom
            sx={{
              fontWeight: 700,
              background: "linear-gradient(45deg, #1976d2, #42a5f5)",
              backgroundClip: "text",
              WebkitBackgroundClip: "text",
              color: "transparent",
            }}
          >
            {t.title}
          </Typography>
          <Typography variant="subtitle1" color="text.secondary">
            {t.subtitle}
          </Typography>
        </Box>

        {/* Main Content */}
        <Paper
          elevation={3}
          sx={{
            p: 3,
            borderRadius: 6,
            background: darkMode
              ? "linear-gradient(to bottom, #1a1a1a, #2d2d2d)"
              : "linear-gradient(to bottom, #ffffff, #f8f9fa)",
          }}
        >
          <Stack spacing={2}>
            {/* File Upload Section */}
            <Box>
              <Typography
                variant="h6"
                gutterBottom
                sx={{ display: "flex", alignItems: "center", gap: 1 }}
              >
                <Typography
                  variant="body2"
                  sx={{
                    bgcolor: "primary.main",
                    color: "primary.contrastText",
                    px: 1,
                    py: 0.5,
                    borderRadius: 1,
                    fontWeight: 500,
                  }}
                >
                  {t.step1}
                </Typography>
                <UploadFileIcon color="primary" />
                {t.fileSelection}
              </Typography>
              <Divider sx={{ mb: 2 }} />
              <Box
                sx={{
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "center",
                }}
              >
                <input
                  accept=".csv"
                  style={{ display: "none" }}
                  id="csv-file-input"
                  multiple
                  type="file"
                  onChange={(e) => setFiles(e.target.files)}
                />
                <label htmlFor="csv-file-input">
                  <Button
                    variant="outlined"
                    component="span"
                    startIcon={<UploadFileIcon />}
                    sx={{
                      minWidth: 200,
                      mb: 2,
                      borderRadius: 2,
                      borderWidth: 2,
                      "&:hover": {
                        borderWidth: 2,
                      },
                    }}
                  >
                    {t.selectFiles}
                  </Button>
                </label>
                {files && (
                  <Typography
                    variant="body2"
                    color="primary"
                    sx={{
                      p: 1,
                      px: 2,
                      borderRadius: 1,
                      bgcolor: "primary.light",
                      color: "primary.contrastText",
                    }}
                  >
                    {files.length} {t.filesSelected}
                  </Typography>
                )}
              </Box>
            </Box>

            {/* Email Section */}
            <Box>
              <Typography
                variant="h6"
                gutterBottom
                sx={{ display: "flex", alignItems: "center", gap: 1 }}
              >
                <Typography
                  variant="body2"
                  sx={{
                    bgcolor: "primary.main",
                    color: "primary.contrastText",
                    px: 1,
                    py: 0.5,
                    borderRadius: 1,
                    fontWeight: 500,
                  }}
                >
                  {t.step2}
                </Typography>
                <EmailIcon color="primary" />
                {t.contactInfo}
              </Typography>
              <Divider sx={{ mb: 2 }} />
              <TextField
                fullWidth
                label={t.emailPrefix}
                value={emailPrefix}
                onChange={(e) => setEmailPrefix(e.target.value)}
                placeholder={t.emailPlaceholder}
                helperText={t.emailHelperText}
                sx={{
                  "& .MuiOutlinedInput-root": {
                    borderRadius: 2,
                  },
                }}
              />
            </Box>

            {/* Advanced Options Section */}
            <Box>
              <Box
                sx={{
                  display: "flex",
                  alignItems: "center",
                  mb: 1,
                  borderRadius: 2,
                  transition: "background-color 0.3s",
                }}
              >
                <Typography
                  variant="h6"
                  sx={{
                    display: "flex",
                    alignItems: "center",
                    gap: 1,
                  }}
                >
                  <Typography
                    variant="body2"
                    sx={{
                      bgcolor: "primary.main",
                      color: "primary.contrastText",
                      px: 1,
                      py: 0.5,
                      borderRadius: 1,
                      fontWeight: 500,
                    }}
                  >
                    {t.optional}
                  </Typography>
                  <SettingsIcon color={"primary"} />
                  {t.advancedOptions}
                </Typography>
                <Box sx={{ flexGrow: 1 }} />
                <Tooltip
                  title={
                    showAdvanced
                      ? "Hide advanced options"
                      : "Show advanced options"
                  }
                >
                  <IconButton
                    onClick={() => setShowAdvanced(!showAdvanced)}
                    sx={{
                      transform: showAdvanced
                        ? "rotate(180deg)"
                        : "rotate(0deg)",
                      transition: "transform 0.3s",
                      color: "primary.main",
                    }}
                  >
                    <ExpandMoreIcon />
                  </IconButton>
                </Tooltip>
              </Box>
              <Collapse in={showAdvanced}>
                <Paper
                  // variant="outlined"
                  sx={{
                    pt: 2,
                    borderRadius: 2,
                    // borderColor: "primary.light",
                    backgroundColor: "#00000000",
                  }}
                  elevation={0}
                >
                  <TextField
                    fullWidth
                    type="number"
                    label={t.rowsPerPage}
                    value={rowsPerPage}
                    onChange={(e) => {
                      const value = parseInt(e.target.value);
                      if (value > 0) {
                        setRowsPerPage(value);
                      }
                    }}
                    inputProps={{ min: 1 }}
                    helperText={t.rowsHelperText}
                    sx={{
                      "& .MuiOutlinedInput-root": {
                        borderRadius: 2,
                      },
                    }}
                  />
                </Paper>
              </Collapse>
            </Box>

            {/* Process Button */}
            <Box>
              <Button
                variant="contained"
                color="primary"
                onClick={processFiles}
                disabled={loading}
                fullWidth
                size="large"
                sx={{
                  py: 1.5,
                  borderRadius: 2,
                  boxShadow: 4,
                  "&:hover": {
                    boxShadow: 6,
                  },
                }}
              >
                {loading ? (
                  <Box sx={{ display: "flex", alignItems: "center", gap: 2 }}>
                    <CircularProgress size={24} color="inherit" />
                    {t.processing}
                  </Box>
                ) : (
                  t.processFiles
                )}
              </Button>
            </Box>
          </Stack>
        </Paper>

        {/* Alerts */}
        <Box sx={{ position: "relative" }}>
          {error && (
            <Alert
              severity="error"
              sx={{
                mb: 2,
                borderRadius: 2,
                boxShadow: 2,
              }}
            >
              {error}
            </Alert>
          )}
          {success && (
            <Alert
              severity="success"
              sx={{
                mb: 2,
                borderRadius: 2,
                boxShadow: 2,
              }}
            >
              {success}
            </Alert>
          )}
        </Box>
      </Stack>
    </Container>
  );
}

export default App;
