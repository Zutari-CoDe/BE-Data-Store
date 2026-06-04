import React, { useState, useRef } from 'react';
import {
  Container,
  TextField,
  Button,
  Card,
  CardContent,
  Grid,
  Typography,
  Box,
  Alert,
  CircularProgress,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Checkbox,
  Paper,
  IconButton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Chip,
  Divider,
  MenuItem,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import DownloadIcon from '@mui/icons-material/Download';
import DeleteIcon from '@mui/icons-material/Delete';
import FolderOpenIcon from '@mui/icons-material/FolderOpen';
import CheckCircleIcon from '@mui/icons-material/CheckCircle';
import AddIcon from '@mui/icons-material/Add';
import ExcelJS from 'exceljs';

const PRIMARY_TEAL = '#00796b';
const LIGHT_TEAL = '#00897b';

export default function ExcelBrandingTool() {
  const [templates, setTemplates] = useState([]);
  const [selectedTemplate, setSelectedTemplate] = useState(null);
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [selectedFiles, setSelectedFiles] = useState(new Set());
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState('info');
  const [templateDialogOpen, setTemplateDialogOpen] = useState(false);
  const [newTemplate, setNewTemplate] = useState({
    name: '',
    companyName: 'Zutari',
    companyInfo: 'Transforming Infrastructure',
    logoUrl: 'Z_symbol.png',
    logoPosition: 'top-right',
    templateFile: null,
  });
  const [docTypeFilter, setDocTypeFilter] = useState('');
  const [disciplineFilter, setDisciplineFilter] = useState('');
  const [saveFolder, setSaveFolder] = useState('');
  const [directoryHandle, setDirectoryHandle] = useState(null);
  const folderInputRef = useRef(null);

  const docTypes = ['Report', 'Specification', 'Design', 'Analysis', 'Plan', 'Other'];
  const disciplines = [
    'Advisory',
    'Architectural',
    'Bridges',
    'Building technology',
    'Civil',
    'Digital & ICT',
    'Electrical',
    'Enviromental and Climate change',
    'Fire Protection',
    'Geotechnical',
    'Hydraulics',
    'Management',
    'Mechanical',
    'Railway',
    'Road and highway',
    'Structural',
    'Transport',
    'Tunneling',
    'Water',
  ];

  const showMessage = (msg, type = 'info') => {
    setMessage(msg);
    setMessageType(type);
    setTimeout(() => setMessage(''), 4000);
  };

  const handleSelectSaveFolder = async () => {
    if (window.showDirectoryPicker) {
      try {
        const handle = await window.showDirectoryPicker();
        setDirectoryHandle(handle);
        setSaveFolder(handle.name || 'Selected folder');
        showMessage('Folder selected. Updated files will be saved there when supported.', 'success');
      } catch (error) {
        if (error.name !== 'AbortError') {
          console.error('Folder selection failed', error);
          showMessage('Unable to select folder. Please try again.', 'error');
        }
      }
      return;
    }

    if (folderInputRef.current) {
      folderInputRef.current.click();
      return;
    }

    showMessage('Folder picker is not supported by this browser.', 'error');
  };

  const handleFolderInputChange = (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    const relativePath = files[0].webkitRelativePath || files[0].name;
    const folderName = relativePath.split('/')[0];
    setSaveFolder(folderName || 'Selected folder');
    setDirectoryHandle(null);
    e.target.value = null;
    showMessage('Folder selected via file picker fallback.', 'success');
  };

  const handleFileUpload = (e) => {
    const files = Array.from(e.target.files || []);
    const newFiles = files.map((file) => ({
      id: Math.random(),
      name: file.name,
      size: (file.size / 1024).toFixed(2),
      docType: docTypeFilter || 'Report',
      discipline: disciplineFilter || 'Civil',
      file: file,
      status: 'pending',
    }));
    setUploadedFiles([...uploadedFiles, ...newFiles]);
    showMessage(`✓ Added ${files.length} file(s)`, 'success');
  };

  const handleFileSelect = (fileId) => {
    const newSelected = new Set(selectedFiles);
    if (newSelected.has(fileId)) {
      newSelected.delete(fileId);
    } else {
      newSelected.add(fileId);
    }
    setSelectedFiles(newSelected);
  };

  const handleSelectAllFiles = (e) => {
    if (e.target.checked) {
      setSelectedFiles(new Set(uploadedFiles.map((f) => f.id)));
    } else {
      setSelectedFiles(new Set());
    }
  };

  const handleUpdateFile = (fileId, field, value) => {
    setUploadedFiles((prev) =>
      prev.map((f) => (f.id === fileId ? { ...f, [field]: value } : f))
    );
  };

  const handleDeleteFile = (fileId) => {
    setUploadedFiles(uploadedFiles.filter((f) => f.id !== fileId));
    const newSelected = new Set(selectedFiles);
    newSelected.delete(fileId);
    setSelectedFiles(newSelected);
  };

  const handleTemplateFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (file) {
      setNewTemplate({ ...newTemplate, templateFile: file });
    }
  };

  const handleSaveTemplate = async () => {
    if (!newTemplate.name || !newTemplate.templateFile) {
      showMessage('Please enter template name and upload a file', 'error');
      return;
    }

    try {
      const arrayBuffer = await newTemplate.templateFile.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const template = {
        id: Math.random(),
        name: newTemplate.name,
        companyName: newTemplate.companyName,
        companyInfo: newTemplate.companyInfo,
        logoUrl: newTemplate.logoUrl,
        logoPosition: newTemplate.logoPosition,
        workbookData: arrayBuffer,
        createdAt: new Date().toLocaleDateString(),
      };

      setTemplates([...templates, template]);
      setTemplateDialogOpen(false);
      setNewTemplate({
        name: '',
        companyName: 'Zutari',
        companyInfo: 'Transforming Infrastructure',
        logoUrl: 'logo.png',
        logoPosition: 'top-right',
        templateFile: null,
      });
      showMessage('✓ Template saved successfully', 'success');
    } catch (error) {
      showMessage(`Error saving template: ${error.message}`, 'error');
    }
  };

  const clearWorksheet = (worksheet) => {
    if (worksheet.rowCount > 0) {
      worksheet.spliceRows(1, worksheet.rowCount);
    }
    if (worksheet.columnCount > 0) {
      worksheet.spliceColumns(1, worksheet.columnCount);
    }
    if (worksheet.model && Array.isArray(worksheet.model.merges)) {
      worksheet.model.merges = [];
    }
    worksheet.pageSetup = {};
    worksheet.headerFooter = {};
    worksheet.printOptions = {};
  };

  const copyWorksheetFromTemplate = (sourceSheet, targetSheet) => {
    clearWorksheet(targetSheet);

    if (sourceSheet.pageSetup) {
      targetSheet.pageSetup = { ...sourceSheet.pageSetup };
    }
    if (sourceSheet.headerFooter) {
      targetSheet.headerFooter = { ...sourceSheet.headerFooter };
    }
    if (sourceSheet.printOptions) {
      targetSheet.printOptions = { ...sourceSheet.printOptions };
    }

    sourceSheet.columns.forEach((sourceCol, index) => {
      const targetCol = targetSheet.getColumn(index + 1);
      targetCol.width = sourceCol.width;
      targetCol.hidden = sourceCol.hidden;
      targetCol.outlineLevel = sourceCol.outlineLevel;
      if (sourceCol.style) {
        targetCol.style = { ...sourceCol.style };
      }
    });

    sourceSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const targetRow = targetSheet.getRow(rowNumber);
      targetRow.height = row.height;
      targetRow.hidden = row.hidden;
      targetRow.outlineLevel = row.outlineLevel;

      for (let col = 1; col <= sourceSheet.columnCount; col += 1) {
        const sourceCell = row.getCell(col);
        const targetCell = targetRow.getCell(col);

        targetCell.value = sourceCell.value;
        if (sourceCell.style) targetCell.style = { ...sourceCell.style };
        if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
        if (sourceCell.alignment) targetCell.alignment = { ...sourceCell.alignment };
        if (sourceCell.border) targetCell.border = { ...sourceCell.border };
        if (sourceCell.fill) targetCell.fill = { ...sourceCell.fill };
        if (sourceCell.protection) targetCell.protection = { ...sourceCell.protection };
        if (sourceCell.hyperlink) targetCell.hyperlink = sourceCell.hyperlink;
        if (sourceCell.note) targetCell.note = sourceCell.note;
      }
    });

    if (sourceSheet.model && Array.isArray(sourceSheet.model.merges)) {
      sourceSheet.model.merges.forEach((range) => {
        targetSheet.mergeCells(range);
      });
    }
  };

  const handleApplyTemplate = async () => {
    if (!selectedTemplate) {
      showMessage('Please select a template', 'error');
      return;
    }

    if (selectedFiles.size === 0) {
      showMessage('Please select at least one file', 'error');
      return;
    }

    setLoading(true);
    try {
      const templateWorkbook = new ExcelJS.Workbook();
      await templateWorkbook.xlsx.load(selectedTemplate.workbookData);
      const templateSheet = templateWorkbook.getWorksheet('DST');

      if (!templateSheet) {
        showMessage('The selected template must contain a sheet named DST', 'error');
        setLoading(false);
        return;
      }

      const filesToProcess = uploadedFiles.filter((f) => selectedFiles.has(f.id));
      let processedCount = 0;
      let skippedCount = 0;

      for (const fileObj of filesToProcess) {
        const userArrayBuffer = await fileObj.file.arrayBuffer();
        const userWorkbook = new ExcelJS.Workbook();
        await userWorkbook.xlsx.load(userArrayBuffer);
        const userSheet = userWorkbook.getWorksheet('DST');

        if (!userSheet) {
          skippedCount += 1;
          setUploadedFiles((prev) =>
            prev.map((f) =>
              f.id === fileObj.id ? { ...f, status: 'skipped' } : f
            )
          );
          continue;
        }

        copyWorksheetFromTemplate(templateSheet, userSheet);

        const excelBuffer = await userWorkbook.xlsx.writeBuffer();
        const blob = new Blob([excelBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const downloadName = `${selectedTemplate.name.replace(/\s+/g, '_')}_${fileObj.name}`;

        if (directoryHandle && directoryHandle.getFileHandle) {
          try {
            const fileHandle = await directoryHandle.getFileHandle(downloadName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(blob);
            await writable.close();
          } catch (error) {
            console.warn('Folder write failed, falling back to browser download', error);
            link.download = downloadName;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
          }
        } else {
          link.download = downloadName;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
        }

        URL.revokeObjectURL(url);

        processedCount += 1;
        setUploadedFiles((prev) =>
          prev.map((f) =>
            f.id === fileObj.id ? { ...f, status: 'completed' } : f
          )
        );
      }

      const summaryParts = [];
      if (processedCount) summaryParts.push(`applied to ${processedCount} file(s)`);
      if (skippedCount) summaryParts.push(`skipped ${skippedCount} file(s) without DST sheet`);
      showMessage(`✓ Template ${summaryParts.join(' and ')}`, 'success');
    } catch (error) {
      console.error('Error applying template:', error);
      showMessage(`Error: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteTemplate = (templateId) => {
    setTemplates(templates.filter((t) => t.id !== templateId));
    if (selectedTemplate?.id === templateId) {
      setSelectedTemplate(null);
    }
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      {message && (
        <Alert severity={messageType} sx={{ mb: 3, borderRadius: 1 }}>
          {message}
        </Alert>
      )}

      {/* Section 1: Templates */}
      <Card
        elevation={0}
        sx={{
          backgroundColor: '#fff',
          border: '1px solid #e0e0e0',
          borderRadius: 2,
          mb: 3,
        }}
      >
        <CardContent sx={{ p: 3 }}>
          <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
            <Typography variant="h6" sx={{ fontWeight: 600, color: PRIMARY_TEAL }}>
              Templates
            </Typography>
            <Button
              variant="contained"
              size="small"
              startIcon={<AddIcon />}
              sx={{
                backgroundColor: PRIMARY_TEAL,
                '&:hover': { backgroundColor: LIGHT_TEAL },
                textTransform: 'none',
              }}
              onClick={() => setTemplateDialogOpen(true)}
            >
              New Template
            </Button>
          </Box>

          {templates.length === 0 ? (
            <Paper sx={{ p: 2, textAlign: 'center', backgroundColor: '#f5f5f5', border: '1px dashed #ccc' }}>
              <Typography variant="body2" color="textSecondary">
                No templates created yet
              </Typography>
            </Paper>
          ) : (
            <Grid container spacing={2}>
              {templates.map((template) => (
                <Grid item xs={12} sm={6} md={4} key={template.id}>
                  <Paper
                    sx={{
                      p: 2,
                      cursor: 'pointer',
                      border: selectedTemplate?.id === template.id
                        ? `2px solid ${PRIMARY_TEAL}`
                        : '1px solid #e0e0e0',
                      backgroundColor: selectedTemplate?.id === template.id ? '#f0f7f6' : '#fff',
                      transition: 'all 0.2s',
                      '&:hover': { boxShadow: 1, borderColor: PRIMARY_TEAL },
                    }}
                    onClick={() => setSelectedTemplate(template)}
                  >
                    <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start' }}>
                      <Box>
                        <Typography variant="subtitle2" sx={{ fontWeight: 600, color: PRIMARY_TEAL }}>
                          {template.name}
                        </Typography>
                        <Typography variant="caption" color="textSecondary">
                          {template.companyName}
                        </Typography>
                      </Box>
                      {selectedTemplate?.id === template.id && (
                        <CheckCircleIcon sx={{ fontSize: 18, color: PRIMARY_TEAL }} />
                      )}
                    </Box>
                    <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mt: 1.5 }}>
                      <Typography variant="caption" color="textSecondary">
                        {template.createdAt}
                      </Typography>
                      <IconButton
                        size="small"
                        sx={{ color: '#d32f2f' }}
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDeleteTemplate(template.id);
                        }}
                      >
                        <DeleteIcon fontSize="small" />
                      </IconButton>
                    </Box>
                  </Paper>
                </Grid>
              ))}
            </Grid>
          )}
        </CardContent>
      </Card>

      <Divider sx={{ my: 3 }} />

      {/* Section 2: Upload Files */}
      <Card
        elevation={0}
        sx={{
          backgroundColor: '#fff',
          border: '1px solid #e0e0e0',
          borderRadius: 2,
          mb: 3,
        }}
      >
        <CardContent sx={{ p: 3 }}>
          <Typography variant="h6" sx={{ fontWeight: 600, color: PRIMARY_TEAL, mb: 2 }}>
            Upload Files
          </Typography>

          <Grid container spacing={2} sx={{ mb: 3 }}>
            <Grid item xs={12} sm={4}>
              <TextField
                select
                fullWidth
                size="small"
                label="Document Type"
                value={docTypeFilter}
                onChange={(e) => setDocTypeFilter(e.target.value)}
                variant="outlined"
                sx={{ minWidth: 200 }}
              >
                <MenuItem value="">All Types</MenuItem>
                {docTypes.map((type) => (
                  <MenuItem key={type} value={type}>
                    {type}
                  </MenuItem>
                ))}
              </TextField>
            </Grid>
            <Grid item xs={12} sm={4}>
              <TextField
                select
                fullWidth
                size="small"
                label="Discipline"
                value={disciplineFilter}
                onChange={(e) => setDisciplineFilter(e.target.value)}
                variant="outlined"
                sx={{ minWidth: 200 }}
              >
                <MenuItem value="">All Disciplines</MenuItem>
                {disciplines.map((disc) => (
                  <MenuItem key={disc} value={disc}>
                    {disc}
                  </MenuItem>
                ))}
              </TextField>
            </Grid>
            <Grid item xs={12} sm={4}>
              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
                <Typography variant="caption" sx={{ color: '#666', fontWeight: 600 }}>
                  Save folder
                </Typography>
                <Button
                  variant="contained"
                  size="small"
                  startIcon={<FolderOpenIcon />}
                  onClick={handleSelectSaveFolder}
                  sx={{ textTransform: 'none', minWidth: 140, width: '100%' }}
                >
                  Choose folder
                </Button>
                <Typography variant="body2" sx={{ color: '#333', minHeight: '1.5rem' }}>
                  {saveFolder || 'No folder selected'}
                </Typography>
                <Typography variant="caption" color="textSecondary">
                  Select a folder for direct output when supported by the browser.
                </Typography>
              </Box>
            </Grid>
          </Grid>

          <input
            ref={folderInputRef}
            type="file"
            webkitdirectory="true"
            directory="true"
            multiple
            style={{ display: 'none' }}
            onChange={handleFolderInputChange}
          />

          <Box
            sx={{
              border: `2px dashed ${PRIMARY_TEAL}`,
              borderRadius: 2,
              p: 4,
              textAlign: 'center',
              backgroundColor: '#f0f7f6',
              cursor: 'pointer',
              transition: 'all 0.2s',
              '&:hover': { backgroundColor: '#e0f2f1', borderColor: LIGHT_TEAL },
              mb: 3,
            }}
          >
            <input
              type="file"
              multiple
              accept=".xlsx,.xls"
              onChange={handleFileUpload}
              style={{ display: 'none' }}
              id="file-input"
            />
            <label htmlFor="file-input" style={{ cursor: 'pointer', display: 'block' }}>
              <CloudUploadIcon sx={{ fontSize: 44, color: PRIMARY_TEAL, mb: 1 }} />
              <Typography variant="body2" sx={{ fontWeight: 600, color: PRIMARY_TEAL, mb: 0.5 }}>
                Click to upload
              </Typography>
              <Typography variant="caption" color="textSecondary">
                or drag and drop Excel files
              </Typography>
            </label>
          </Box>

          {uploadedFiles.length > 0 && (
            <Box>
              <Typography variant="subtitle2" sx={{ fontWeight: 600, color: '#222', mb: 2 }}>
                Files ({uploadedFiles.length})
              </Typography>
              <TableContainer sx={{ border: '1px solid #e0e0e0', borderRadius: 1 }}>
                <Table size="small">
                  <TableHead>
                    <TableRow sx={{ backgroundColor: '#fafafa', borderBottom: '2px solid #e0e0e0' }}>
                      <TableCell padding="checkbox">
                        <Checkbox
                          size="small"
                          indeterminate={
                            selectedFiles.size > 0 &&
                            selectedFiles.size < uploadedFiles.length
                          }
                          checked={
                            uploadedFiles.length > 0 &&
                            selectedFiles.size === uploadedFiles.length
                          }
                          onChange={handleSelectAllFiles}
                        />
                      </TableCell>
                      <TableCell sx={{ fontWeight: 600, color: PRIMARY_TEAL, fontSize: '13px' }}>
                        File Name
                      </TableCell>
                      <TableCell sx={{ fontWeight: 600, color: PRIMARY_TEAL, fontSize: '13px' }}>
                        Doc Type
                      </TableCell>
                      <TableCell sx={{ fontWeight: 600, color: PRIMARY_TEAL, fontSize: '13px' }}>
                        Discipline
                      </TableCell>
                      <TableCell sx={{ fontWeight: 600, color: PRIMARY_TEAL, fontSize: '13px' }}>
                        Size
                      </TableCell>
                      <TableCell sx={{ fontWeight: 600, color: PRIMARY_TEAL, fontSize: '13px' }}>
                        Status
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600, color: PRIMARY_TEAL, fontSize: '13px' }}>
                        Action
                      </TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {uploadedFiles.map((file) => (
                      <TableRow
                        key={file.id}
                        sx={{
                          backgroundColor: selectedFiles.has(file.id) ? '#f0f7f6' : '#fff',
                          '&:hover': { backgroundColor: '#fafafa' },
                          borderBottom: '1px solid #e0e0e0',
                        }}
                      >
                        <TableCell padding="checkbox">
                          <Checkbox
                            size="small"
                            checked={selectedFiles.has(file.id)}
                            onChange={() => handleFileSelect(file.id)}
                          />
                        </TableCell>
                        <TableCell sx={{ fontSize: '13px' }}>{file.name}</TableCell>
                        <TableCell>
                          <TextField
                            select
                            size="small"
                            value={file.docType}
                            onChange={(e) => handleUpdateFile(file.id, 'docType', e.target.value)}
                            variant="outlined"
                            sx={{ width: 130 }}
                          >
                            {docTypes.map((type) => (
                              <MenuItem key={type} value={type}>
                                {type}
                              </MenuItem>
                            ))}
                          </TextField>
                        </TableCell>
                        <TableCell>
                          <TextField
                            select
                            size="small"
                            value={file.discipline}
                            onChange={(e) => handleUpdateFile(file.id, 'discipline', e.target.value)}
                            variant="outlined"
                            sx={{ width: 130 }}
                          >
                            {disciplines.map((disc) => (
                              <MenuItem key={disc} value={disc}>
                                {disc}
                              </MenuItem>
                            ))}
                          </TextField>
                        </TableCell>
                        <TableCell sx={{ fontSize: '13px' }}>{file.size} KB</TableCell>
                        <TableCell>
                          <Chip
                            label={file.status === 'completed' ? '✓ Done' : 'Pending'}
                            size="small"
                            sx={{
                              backgroundColor: file.status === 'completed' ? '#c8e6c9' : '#fff3cd',
                              color: file.status === 'completed' ? '#2e7d32' : '#856404',
                              fontWeight: 500,
                              fontSize: '12px',
                            }}
                          />
                        </TableCell>
                        <TableCell align="right">
                          <IconButton
                            size="small"
                            sx={{ color: '#d32f2f' }}
                            onClick={() => handleDeleteFile(file.id)}
                          >
                            <DeleteIcon fontSize="small" />
                          </IconButton>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            </Box>
          )}
        </CardContent>
      </Card>

      <Divider sx={{ my: 3 }} />

      {/* Section 3: Apply Template */}
      <Card
        elevation={0}
        sx={{
          backgroundColor: '#fff',
          border: '1px solid #e0e0e0',
          borderRadius: 2,
        }}
      >
        <CardContent sx={{ p: 3 }}>
          <Typography variant="h6" sx={{ fontWeight: 600, color: PRIMARY_TEAL, mb: 2 }}>
            Apply Template
          </Typography>

          <Paper
            sx={{
              p: 2.5,
              backgroundColor: '#f0f7f6',
              border: `1px solid ${PRIMARY_TEAL}`,
              borderRadius: 1,
              mb: 3,
            }}
          >
            <Grid container spacing={2}>
              <Grid item xs={12} sm={6}>
                <Typography variant="caption" color="textSecondary" sx={{ fontWeight: 600 }}>
                  SELECTED FILES
                </Typography>
                <Typography variant="h6" sx={{ fontWeight: 600, color: PRIMARY_TEAL }}>
                  {selectedFiles.size} of {uploadedFiles.length}
                </Typography>
              </Grid>
              <Grid item xs={12} sm={6}>
                <Typography variant="caption" color="textSecondary" sx={{ fontWeight: 600 }}>
                  SELECTED TEMPLATE
                </Typography>
                <Typography variant="h6" sx={{ fontWeight: 600, color: PRIMARY_TEAL }}>
                  {selectedTemplate?.name || '—'}
                </Typography>
              </Grid>
            </Grid>
          </Paper>

          <Typography variant="body2" color="textSecondary" sx={{ mb: 2 }}>
            The selected template's DST sheet will replace the DST sheet in each selected Excel file.
          </Typography>

          <Box sx={{ display: 'flex', justifyContent: 'center' }}>
            <Button
              variant="contained"
              size="large"
              sx={{
                backgroundColor: PRIMARY_TEAL,
                '&:hover': { backgroundColor: LIGHT_TEAL },
                textTransform: 'none',
                fontWeight: 600,
                px: 5,
                py: 1.5,
              }}
              startIcon={loading ? <CircularProgress size={20} color="inherit" /> : <DownloadIcon />}
              onClick={handleApplyTemplate}
              disabled={!selectedTemplate || selectedFiles.size === 0 || loading}
            >
              {loading ? 'Processing...' : 'Apply & Download'}
            </Button>
          </Box>
        </CardContent>
      </Card>

      {/* Template Dialog */}
      <Dialog open={templateDialogOpen} onClose={() => setTemplateDialogOpen(false)} maxWidth="sm" fullWidth>
        <DialogTitle
          sx={{
            backgroundColor: PRIMARY_TEAL,
            color: 'white',
            fontWeight: 600,
            fontSize: '16px',
          }}
        >
          Create New Template
        </DialogTitle>
        <DialogContent sx={{ pt: 3 }}>
          <TextField
            fullWidth
            label="Template Name"
            value={newTemplate.name}
            onChange={(e) => setNewTemplate({ ...newTemplate, name: e.target.value })}
            margin="normal"
            variant="outlined"
            size="small"
          />
          <TextField
            fullWidth
            label="Company Name"
            value={newTemplate.companyName}
            onChange={(e) => setNewTemplate({ ...newTemplate, companyName: e.target.value })}
            margin="normal"
            variant="outlined"
            size="small"
          />
          <TextField
            fullWidth
            label="Company Info"
            value={newTemplate.companyInfo}
            onChange={(e) => setNewTemplate({ ...newTemplate, companyInfo: e.target.value })}
            margin="normal"
            variant="outlined"
            size="small"
            multiline
            rows={2}
          />
          <TextField
            fullWidth
            label="Logo File (relative path or URL)"
            value={newTemplate.logoUrl}
            onChange={(e) => setNewTemplate({ ...newTemplate, logoUrl: e.target.value })}
            margin="normal"
            variant="outlined"
            size="small"
            placeholder="Z_symbol.png or https://example.com/logo.png"
            helperText='Place logo in "public" folder or provide full URL'
          />
          <Box sx={{ mt: 2 }}>
            <Typography variant="caption" sx={{ fontWeight: 600, color: PRIMARY_TEAL, display: 'block', mb: 1 }}>
              TEMPLATE FILE
            </Typography>
            <Box
              sx={{
                border: `1px dashed ${PRIMARY_TEAL}`,
                borderRadius: 1,
                p: 2,
                textAlign: 'center',
                backgroundColor: '#f0f7f6',
              }}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleTemplateFileUpload}
                style={{ display: 'none' }}
                id="template-file-input"
              />
              <label htmlFor="template-file-input" style={{ cursor: 'pointer', display: 'block' }}>
                <CloudUploadIcon sx={{ color: PRIMARY_TEAL, mb: 0.5, fontSize: 28 }} />
                <Typography variant="caption">
                  {newTemplate.templateFile ? `✓ ${newTemplate.templateFile.name}` : 'Click to upload'}
                </Typography>
              </label>
            </Box>
          </Box>
        </DialogContent>
        <DialogActions sx={{ p: 2 }}>
          <Button onClick={() => setTemplateDialogOpen(false)} size="small">
            Cancel
          </Button>
          <Button
            variant="contained"
            sx={{
              backgroundColor: PRIMARY_TEAL,
              '&:hover': { backgroundColor: LIGHT_TEAL },
              textTransform: 'none',
            }}
            onClick={handleSaveTemplate}
            size="small"
          >
            Save
          </Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
