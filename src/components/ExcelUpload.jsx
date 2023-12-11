import { useState } from "react";
import * as ExcelJS from "exceljs";
import "tailwindcss/tailwind.css"; // 請確保這一行已經引入 Tailwind CSS

const ExcelUpload = () => {
  const [file, setFile] = useState(null);
  const [data, setData] = useState([]);

  const handleFileChange = async (e) => {
    const selectedFile = e.target.files[0];

    if (selectedFile) {
      setFile(selectedFile);

      // 使用exceljs讀取文件內容
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(selectedFile);
      const worksheet = workbook.worksheets[0];

      // 將數據轉換為陣列
      const extractedData = [];
      worksheet.eachRow((row) => {
        const rowData = [];
        row.eachCell((cell) => {
          rowData.push(cell.value);
        });
        extractedData.push(rowData);
      });
      // console.log(extractedData);
      setData(extractedData);
    }
  };
  const handleDownload = async () => {
    const bl = new ExcelJS.Workbook();
    const blsheet = bl.addWorksheet("Sheet 1");
    const wpk = new ExcelJS.Workbook();
    const wpksheet = wpk.addWorksheet("Sheet 1");
    const pkw = new ExcelJS.Workbook();
    const pkwsheet = pkw.addWorksheet("Sheet 1");
    const wptg = new ExcelJS.Workbook();
    const wptgsheet = wptg.addWorksheet("Sheet 1");
    // console.log(data);
    data.forEach((arr) => {
      // console.log(arr);
      arr[3] == "buluo"
        ? blsheet.addRow(arr)
        : arr[3] == "wpk"
        ? wpksheet.addRow(arr)
        : arr[3] == "pkw"
        ? pkwsheet.addRow(arr)
        : wptgsheet.addRow(arr);
    });
    const buffer = await bl.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = "excel_data.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className=" p-6">
      <h2 className="text-2xl font-semibold mb-4">上傳並讀取XLSX文件</h2>
      <input
        type="file"
        accept=".xlsx"
        onChange={handleFileChange}
        className="mt-2 p-2 border border-gray-600 rounded-md bg-gray-700 text-gray-200"
      />
      <div className="mt-4">
        <button
          onClick={handleDownload}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載修改後的文件
        </button>
      </div>
      {file && (
        <div className="mt-4">
          <h3 className="text-xl mb-2">已選擇文件: {file.name}</h3>
          {/* <table className="table-auto w-full">
            <thead>
              <tr>
                {data.length > 0 &&
                  data[0].map((header, index) => (
                    <th key={index} className="px-4 py-2">
                      {header}
                    </th>
                  ))}
              </tr>
            </thead>
            <tbody>
              {data.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, cellIndex) => (
                    <td key={cellIndex} className="border px-4 py-2">
                      {cell}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table> */}
        </div>
      )}
    </div>
  );
};

export default ExcelUpload;
