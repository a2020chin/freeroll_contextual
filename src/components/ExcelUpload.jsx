import { useState } from "react";
import * as ExcelJS from "exceljs";

const ExcelUpload = () => {
  const [file, setFile] = useState(null);
  const [month, setMonth] = useState(3);
  const [ipfilter, setIpfilter] = useState(20);
  const [blData, setBlData] = useState([]);
  const [wpkData, setWpkData] = useState([]);
  const [pkwData, setPkwData] = useState([]);
  const [wptgData, setWptgData] = useState([]);

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
      // setData(extractedData);
      const nowDate = new Date();
      await nowDate.setMonth(nowDate.getMonth() - month);

      extractedData.forEach((arr) => {
        if (arr[8] >= ipfilter) {
          if (nowDate <= new Date(arr[4])) {
            arr[3] == "buluo"
              ? setBlData((prevBlData) => [...prevBlData, arr])
              : arr[3] == "wpk"
              ? setWpkData((prevBlData) => [...prevBlData, arr])
              : arr[3] == "pkw"
              ? setPkwData((prevBlData) => [...prevBlData, arr])
              : setWptgData((prevBlData) => [...prevBlData, arr]);
          }
        }
      });
    }
  };
  const download = async (platData, plat) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");
    const platarr = [];

    platData.forEach((arr) => {
      platarr.push(arr[0], arr[13].match(/\d+/g));
    });

    const filterPlat = [...new Set(platarr.flat(Infinity))];
    const filterPlatarr = filterPlat.filter((element) => element !== null);

    filterPlatarr.forEach((arr) => {
      worksheet.addRow(Array(String(arr)));
    });

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = `${plat}.xlsx`;
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
      <div className="flex gap-x-4">
        <div>
          <label htmlFor="monthInput" className="text-white">
            過濾幾個月份：
          </label>
          <input
            type="number"
            id="monthInput"
            value={month}
            onChange={(e) => {
              setMonth(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
        <div>
          <label htmlFor="ipInput" className="text-white">
            過濾幾個ip：
          </label>
          <input
            type="number"
            id="ipInput"
            value={ipfilter}
            onChange={(e) => {
              setIpfilter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
      </div>
      <div className="flex mt-4 gap-2">
        <button
          onClick={() => download(blData, "BL")}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載BL
        </button>
        <button
          onClick={() => download(wpkData, "WPK")}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載WPK
        </button>
        <button
          onClick={() => download(pkwData, "PKW")}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載PKW
        </button>
        <button
          onClick={() => download(wptgData, "WPTG")}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載WPTG
        </button>
      </div>
    </div>
  );
};

export default ExcelUpload;
