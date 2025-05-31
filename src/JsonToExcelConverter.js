import React, { useState, useRef } from 'react';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Upload, FileText, Download, AlertCircle, CheckCircle, X } from 'lucide-react';

export default function JsonToExcelConverter() {
  const [jsonDataList, setJsonDataList] = useState([]);
  const [fileNames, setFileNames] = useState([]);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [preview, setPreview] = useState([]);
  const [totalRows, setTotalRows] = useState(0);
  const [originalJsonString, setOriginalJsonString] = useState('');
  const fileInputRef = useRef(null);

  // Handle multiple file uploads
  const handleFileUpload = (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    // Reset states
    setError('');
    setSuccess('');
    setIsLoading(true);
    
    // Keep track of processed files
    let processedCount = 0;
    const newJsonDataList = [...jsonDataList];
    const newFileNames = [...fileNames];
    
    // Process each file
    Array.from(files).forEach(file => {
      // Check if file is JSON
      if (file.type !== 'application/json' && !file.name.endsWith('.json')) {
        setError(prev => `${prev ? prev + '\n' : ''}${file.name}: Not a valid JSON file`);
        processedCount++;
        
        if (processedCount === files.length) {
          setIsLoading(false);
        }
        return;
      }
      
      // Add file name to the list
      newFileNames.push(file.name.replace('.json', ''));
      
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = JSON.parse(event.target.result);
          newJsonDataList.push(data);
          
          processedCount++;
          
          // When all files are processed
          if (processedCount === files.length) {
            setJsonDataList(newJsonDataList);
            setFileNames(newFileNames);
            
            // Generate preview from combined data
            generatePreview(newJsonDataList);
            setIsLoading(false);
            setSuccess(`${files.length} file(s) loaded successfully`);
          }
        } catch (err) {
          setError(prev => `${prev ? prev + '\n' : ''}${file.name}: Invalid JSON format: ${err.message}`);
          processedCount++;
          
          if (processedCount === files.length) {
            setIsLoading(false);
          }
        }
      };
      
      reader.onerror = () => {
        setError(prev => `${prev ? prev + '\n' : ''}${file.name}: Error reading file`);
        processedCount++;
        
        if (processedCount === files.length) {
          setIsLoading(false);
        }
      };
      
      reader.readAsText(file);
    });
  };

  // Process JSON data from the text area
  const handleJsonInput = (jsonString) => {
    setOriginalJsonString(jsonString);
    
    try {
      if (jsonString.trim()) {
        const data = JSON.parse(jsonString);
        
        // Add to the existing data list
        const newJsonDataList = [...jsonDataList, data];
        setJsonDataList(newJsonDataList);
        
        // Add a generic filename for this input
        setFileNames([...fileNames, `manual_input_${fileNames.length + 1}`]);
        
        // Update preview
        generatePreview(newJsonDataList);
        setError('');
        setSuccess('JSON parsed successfully');
      } else {
        setPreview([]);
      }
    } catch (err) {
      setError('Invalid JSON format: ' + err.message);
    }
  };

  // Extract specific headers and convert to table format
  const extractSpecificHeaders = (data) => {
    // Define all the headers we want to extract
    const headerMap = {
      'SenderID': ['Remittance', 'Header', 'SenderID'],
      'ReceiverID': ['Remittance', 'Header', 'ReceiverID'],
      'TransactionDate': ['Remittance', 'Header', 'TransactionDate'],
      'RecordCount': ['Remittance', 'Header', 'RecordCount'],
      'DispositionFlag': ['Remittance', 'Header', 'DispositionFlag'],
      'PayerID': ['Remittance', 'Header', 'PayerID'],
      'ClaimID': ['Remittance', 'Claim', null, 'ID'],
      'IDPayer': ['Remittance', 'Claim', null, 'IDPayer'],
      'ProviderID': ['Remittance', 'Claim', null, 'ProviderID'],
      'PaymentReference': ['Remittance', 'Claim', null, 'PaymentReference'],
      'DateSettlement': ['Remittance', 'Claim', null, 'DateSettlement'],
      'FacilityID': ['Remittance', 'Claim', null, 'Encounter', 'FacilityID'],
      'ActivityID': ['Remittance', 'Claim', null, 'Activity', null, 'ID'],
      'Start': ['Remittance', 'Claim', null, 'Activity', null, 'Start'],
      'Type': ['Remittance', 'Claim', null, 'Activity', null, 'Type'],
      'Code': ['Remittance', 'Claim', null, 'Activity', null, 'Code'],
      'Quantity': ['Remittance', 'Claim', null, 'Activity', null, 'Quantity'],
      'Net': ['Remittance', 'Claim', null, 'Activity', null, 'Net'],
      'Clinician': ['Remittance', 'Claim', null, 'Activity', null, 'Clinician'],
      'Gross': ['Remittance', 'Claim', null, 'Activity', null, 'Gross'],
      'PatientShare': ['Remittance', 'Claim', null, 'Activity', null, 'PatientShare'],
      'PaymentAmount': ['Remittance', 'Claim', null, 'Activity', null, 'PaymentAmount'],
      'DenialCode': ['Remittance', 'Claim', null, 'Activity', null, 'DenialCode'],
      'Comments': ['Remittance', 'Claim', null, 'Activity', null, 'Comments'],
      'PriorAuthorizationID': ['Remittance', 'Claim', null, 'Activity', null, 'PriorAuthorizationID']
    };
    
    // Function to safely navigate nested properties
    const getNestedValue = (obj, path, claimIndex = null, activityIndex = null) => {
      if (!obj) return null;
      
      let current = obj;
      
      for (let i = 0; i < path.length; i++) {
        let key = path[i];
        
        // Handle special cases for claim and activity indices
        if (key === null && i === 2) {
          key = claimIndex;
        } else if (key === null && i === 4) {
          key = activityIndex;
        }
        
        // Check if the current key exists
        if (current[key] === undefined) {
          return null;
        }
        
        current = current[key];
      }
      
      return current;
    };
    
    // Create a flat table structure
    const tableRows = [];
    
    // Check if we have the expected structure
    if (data && data.Remittance && data.Remittance.Claim && Array.isArray(data.Remittance.Claim)) {
      const claims = data.Remittance.Claim;
      
      // Process each claim
      claims.forEach((claim, claimIndex) => {
        // Check if this claim has activities
        if (claim.Activity && Array.isArray(claim.Activity) && claim.Activity.length > 0) {
          // Create a row for each activity
          claim.Activity.forEach((activity, activityIndex) => {
            const row = {};
            
            // Extract all headers
            for (const [header, path] of Object.entries(headerMap)) {
              row[header] = getNestedValue(data, path, claimIndex, activityIndex);
            }
            
            tableRows.push(row);
          });
        } else {
          // Create a row just for the claim without activities
          const row = {};
          
          // Extract all claim-level headers
          for (const [header, path] of Object.entries(headerMap)) {
            // Skip activity-specific headers
            if (!path.includes('Activity')) {
              row[header] = getNestedValue(data, path, claimIndex);
            }
          }
          
          tableRows.push(row);
        }
      });
    } else {
      // If the structure is different, try a more generic approach
      return fallbackToGenericTable(data);
    }
    
    return tableRows;
  };
  
  // Generic table conversion as a fallback
  const fallbackToGenericTable = (data) => {
    // Handle arrays of objects directly
    if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object') {
      return data;
    }
    
    // For objects, flatten them
    if (typeof data === 'object' && data !== null) {
      const flattenObject = (obj, prefix = '') => {
        const result = {};
        
        for (const key in obj) {
          const newKey = prefix ? `${prefix}.${key}` : key;
          
          if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
            Object.assign(result, flattenObject(obj[key], newKey));
          } else if (Array.isArray(obj[key])) {
            result[newKey] = JSON.stringify(obj[key]);
          } else {
            result[newKey] = obj[key];
          }
        }
        
        return result;
      };
      
      return [flattenObject(data)];
    }
    
    // For primitive values
    return [{ value: data }];
  };

  // Generate preview from multiple JSON files
  const generatePreview = (dataList) => {
    try {
      if (!dataList || dataList.length === 0) {
        setPreview([]);
        setTotalRows(0);
        return;
      }
      
      // Process and combine data from all files
      let allTableRows = [];
      
      dataList.forEach((data, index) => {
        // Extract rows from this data
        const tableData = extractSpecificHeaders(data);
        
        // Add source file information to each row (removed as per user request)
        const rowsWithSource = tableData.map(row => ({
          // Remove SourceFile column
          ...row
        }));
        
        // Add to the combined list
        allTableRows = [...allTableRows, ...rowsWithSource];
      });
      
      // Store total row count
      setTotalRows(allTableRows.length);
      
      // Show preview of combined data
      setPreview(allTableRows.slice(0, 5)); // Show first 5 items
    } catch (err) {
      console.error("Error generating preview:", err);
      setError("Error generating preview: " + err.message);
      setPreview([]);
      setTotalRows(0);
    }
  };

  // Convert JSON to Excel with styled borders using ExcelJS
  const convertToExcel = () => {
    if (!jsonDataList || jsonDataList.length === 0) {
      setError('No data to convert');
      return;
    }

    try {
      // Process and combine data from all files
      let allTableRows = [];
      
      jsonDataList.forEach((data, index) => {
        // Extract rows from this data
        const tableData = extractSpecificHeaders(data);
        
        // Simply add the rows without source file information
        allTableRows = [...allTableRows, ...tableData];
      });
      
      // Create a new workbook
      const workbook = new ExcelJS.Workbook();
      
      // Add a worksheet
      const worksheet = workbook.addWorksheet('Combined Data');
      
      // Add headers
      if (allTableRows.length > 0) {
        const headers = Object.keys(allTableRows[0]);
        
        // Add header row
        worksheet.addRow(headers);
        
        // Style header row
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.height = 22;
        
        // Add all data rows
        allTableRows.forEach(row => {
          const values = headers.map(header => row[header] !== null ? row[header] : '');
          worksheet.addRow(values);
        });
        
        // Auto-fit columns
        headers.forEach((header, i) => {
          const column = worksheet.getColumn(i + 1);
          
          // Find max width
          let maxLength = header.length;
          
          // Sample rows for width
          const sampleSize = Math.min(20, allTableRows.length);
          for (let j = 0; j < sampleSize; j++) {
            const value = String(allTableRows[j][header] || '');
            const valueLength = value.length;
            if (valueLength > maxLength) {
              maxLength = Math.min(valueLength, 50); // Cap at 50 chars
            }
          }
          
          // Set column width with some padding
          column.width = maxLength + 3;
        });
        
        // Apply thin borders to all cells
        const totalRows = allTableRows.length + 1; // +1 for header
        
        // Apply borders to all cells
        for (let rowIndex = 1; rowIndex <= totalRows; rowIndex++) {
          const row = worksheet.getRow(rowIndex);
          
          headers.forEach((_, colIndex) => {
            const cell = row.getCell(colIndex + 1);
            
            // Apply thin borders
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            
            // Center-align header cells
            if (rowIndex === 1) {
              cell.alignment = { 
                horizontal: 'center',
                vertical: 'middle'
              };
              
              // Add light gray background to headers
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF2F2F2' }
              };
            } else {
              // Align data cells
              if (typeof cell.value === 'number') {
                cell.alignment = { horizontal: 'right' };
              } else {
                cell.alignment = { horizontal: 'left' };
              }
            }
          });
        }
        
        // Style alternating rows for better readability
        for (let rowIndex = 2; rowIndex <= totalRows; rowIndex++) {
          if (rowIndex % 2 === 0) {
            const row = worksheet.getRow(rowIndex);
            
            headers.forEach((_, colIndex) => {
              const cell = row.getCell(colIndex + 1);
              
              // Add very light gray background to alternating rows
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFAFAFA' }
              };
            });
          }
        }
        
        // Freeze the header row
        worksheet.views = [
          { state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }
        ];
        
        // Add filters to headers for easy sorting - REMOVE THIS IF YOU DON'T WANT FILTER DROPDOWNS
        /*
        worksheet.autoFilter = {
          from: { row: 1, column: 1 },
          to: { row: 1, column: headers.length }
        };
        */
        
        // Write the file and save it
        workbook.xlsx.writeBuffer().then(buffer => {
          const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          saveAs(blob, `combined_data.xlsx`);
          setSuccess(`Combined data from ${jsonDataList.length} file(s) exported successfully`);
        });
      } else {
        setError('No data to convert');
      }
    } catch (err) {
      setError('Error converting to Excel: ' + err.message);
    }
  };

  // Reset the form
  const resetForm = () => {
    setJsonDataList([]);
    setFileNames([]);
    setError('');
    setSuccess('');
    setPreview([]);
    setTotalRows(0);
    setOriginalJsonString('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-indigo-50 to-white p-4 md:p-8">
      <div className="max-w-4xl mx-auto">
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-indigo-800 mb-2">JSON to Excel Converter</h1>
          <p className="text-gray-600">Upload JSON files and convert them to a single Excel spreadsheet</p>
        </div>

        {/* File Upload Card */}
        <div className="bg-white rounded-xl shadow-lg p-6 mb-8 transition-all">
          <div 
            className="border-2 border-dashed border-indigo-300 rounded-lg p-8 text-center hover:border-indigo-500 transition-colors cursor-pointer"
            onClick={() => fileInputRef.current.click()}
          >
            <input 
              type="file" 
              ref={fileInputRef}
              onChange={handleFileUpload}
              className="hidden" 
              accept=".json"
              multiple
            />
            
            <div className="flex flex-col items-center">
              <Upload className="h-12 w-12 text-indigo-500 mb-3" />
              <h3 className="font-medium text-lg text-gray-800 mb-1">Upload JSON Files</h3>
              <p className="text-gray-500 text-sm mb-4">Click to browse or drag and drop (multiple files supported)</p>
              <button 
                className="bg-indigo-600 text-white px-4 py-2 rounded-md hover:bg-indigo-700 transition-colors inline-flex items-center"
                onClick={(e) => {
                  e.stopPropagation();
                  fileInputRef.current.click();
                }}
              >
                <FileText className="h-4 w-4 mr-2" />
                Select Files
              </button>
            </div>
          </div>
          
          {/* Text area for manual JSON input */}
          <div className="mt-6">
            <h3 className="font-medium text-lg text-gray-800 mb-2">Or paste JSON directly:</h3>
            <textarea
              className="w-full h-40 border border-gray-300 rounded-md p-3 font-mono text-sm"
              placeholder="Paste your JSON here..."
              value={originalJsonString}
              onChange={(e) => {
                setOriginalJsonString(e.target.value);
                handleJsonInput(e.target.value);
              }}
            ></textarea>
          </div>
        </div>

        {/* File List */}
        {fileNames.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg p-6 mb-8">
            <h3 className="text-lg font-semibold text-gray-800 mb-4">Loaded Files ({fileNames.length})</h3>
            <ul className="divide-y divide-gray-200">
              {fileNames.map((name, index) => (
                <li key={index} className="py-2 flex items-center justify-between">
                  <div className="flex items-center">
                    <FileText className="h-5 w-5 text-indigo-500 mr-3" />
                    <span className="text-gray-700">{name}</span>
                  </div>
                  <button 
                    onClick={() => {
                      // Remove this file
                      const newFileNames = [...fileNames];
                      const newJsonDataList = [...jsonDataList];
                      newFileNames.splice(index, 1);
                      newJsonDataList.splice(index, 1);
                      setFileNames(newFileNames);
                      setJsonDataList(newJsonDataList);
                      generatePreview(newJsonDataList);
                    }}
                    className="text-red-500 hover:text-red-700"
                  >
                    <X className="h-4 w-4" />
                  </button>
                </li>
              ))}
            </ul>
          </div>
        )}

        {/* Alerts */}
        {error && (
          <div className="bg-red-50 border-l-4 border-red-500 p-4 mb-6 rounded-md flex items-start">
            <AlertCircle className="h-5 w-5 text-red-500 mr-3 mt-0.5 flex-shrink-0" />
            <div className="flex-grow">
              <p className="text-red-800 font-medium">Error</p>
              <p className="text-red-700 text-sm" style={{ whiteSpace: 'pre-line' }}>{error}</p>
            </div>
            <button onClick={() => setError('')} className="text-red-500 hover:text-red-700">
              <X className="h-5 w-5" />
            </button>
          </div>
        )}

        {success && (
          <div className="bg-green-50 border-l-4 border-green-500 p-4 mb-6 rounded-md flex items-start">
            <CheckCircle className="h-5 w-5 text-green-500 mr-3 mt-0.5 flex-shrink-0" />
            <div className="flex-grow">
              <p className="text-green-800 font-medium">Success</p>
              <p className="text-green-700 text-sm">{success}</p>
            </div>
            <button onClick={() => setSuccess('')} className="text-green-500 hover:text-green-700">
              <X className="h-5 w-5" />
            </button>
          </div>
        )}

        {/* Preview */}
        {preview.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg p-6 mb-8">
            <h3 className="text-lg font-semibold text-gray-800 mb-4">Preview</h3>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    {Object.keys(preview[0]).map((key) => (
                      <th 
                        key={key} 
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider"
                      >
                        {key}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {preview.map((row, idx) => (
                    <tr key={idx} className={idx % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                      {Object.keys(row).map((key) => {
                        let displayValue = row[key];
                        
                        // Handle complex nested objects
                        if (typeof displayValue === 'object' && displayValue !== null) {
                          displayValue = JSON.stringify(displayValue).substring(0, 50) + 
                            (JSON.stringify(displayValue).length > 50 ? '...' : '');
                        }
                        
                        return (
                          <td key={key} className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                            {String(displayValue !== null ? displayValue : '')}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.length < totalRows && (
                <p className="text-sm text-gray-500 mt-3">
                  Showing {preview.length} of {totalRows} records
                </p>
              )}
            </div>
          </div>
        )}

        {/* Action Buttons */}
        <div className="flex flex-col sm:flex-row gap-4 justify-center">
          {jsonDataList.length > 0 && (
            <>
              <button 
                onClick={convertToExcel} 
                className="bg-green-600 hover:bg-green-700 text-white px-6 py-3 rounded-lg shadow-md transition-colors flex items-center justify-center"
                disabled={isLoading}
              >
                <Download className="h-5 w-5 mr-2" />
                Convert to Excel
              </button>
              
              <button 
                onClick={resetForm} 
                className="bg-gray-600 hover:bg-gray-700 text-white px-6 py-3 rounded-lg shadow-md transition-colors"
                disabled={isLoading}
              >
                Start Over
              </button>
            </>
          )}
        </div>

        {/* Loading indicator */}
        {isLoading && (
          <div className="flex justify-center mt-6">
            <div className="w-12 h-12 border-4 border-indigo-200 border-t-indigo-600 rounded-full animate-spin"></div>
          </div>
        )}
      </div>
    </div>
  );
}