import { useState } from "react";
import * as ExcelJS from "exceljs";
import "tailwindcss/tailwind.css"; // 請確保這一行已經引入 Tailwind CSS

const ExcelUpload = () => {
  const [file, setFile] = useState(null);
  // const [data, setData] = useState([]);
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
      await nowDate.setMonth(nowDate.getMonth() - 6);

      extractedData.forEach((arr) => {
        if (nowDate <= new Date(arr[4])) {
          arr[3] == "buluo"
            ? setBlData((prevBlData) => [...prevBlData, arr])
            : arr[3] == "wpk"
            ? setWpkData((prevBlData) => [...prevBlData, arr])
            : arr[3] == "pkw"
            ? setPkwData((prevBlData) => [...prevBlData, arr])
            : setWptgData((prevBlData) => [...prevBlData, arr]);
        }
      });
    }
  };
  const blDownload = async () => {
    const bl = new ExcelJS.Workbook();
    const blsheet = bl.addWorksheet("Sheet 1");
    const blarr = [];

    blData.forEach((arr) => {
      blarr.push(arr[0], arr[13].match(/\d+/g));
    });

    const filterBl = [...new Set(blarr.flat(Infinity))];
    const filterBlarr = filterBl.filter((element) => element !== null);

    filterBlarr.forEach((arr) => {
      blsheet.addRow(Array(String(arr)));
    });

    const buffer = await bl.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = "BL.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const wpkDownload = async () => {
    const wpk = new ExcelJS.Workbook();
    const wpksheet = wpk.addWorksheet("Sheet 1");
    const wpkarr = [];

    wpkData.forEach((arr) => {
      wpkarr.push(arr[0], arr[13].match(/\d+/g));
    });

    const filterWpk = [...new Set(wpkarr.flat(Infinity))];
    const filterWpkarr = filterWpk.filter((element) => element !== null);

    filterWpkarr.forEach((arr) => {
      wpksheet.addRow(Array(String(arr)));
    });

    const buffer = await wpk.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = "WPK.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
  const pkwDownload = async () => {
    const pkw = new ExcelJS.Workbook();
    const pkwsheet = pkw.addWorksheet("Sheet 1");
    const pkwarr = [];

    pkwData.forEach((arr) => {
      pkwarr.push(arr[0], arr[13].match(/\d+/g));
    });

    const filterPkw = [...new Set(pkwarr.flat(Infinity))];
    const filterPkwarr = filterPkw.filter((element) => element !== null);

    filterPkwarr.forEach((arr) => {
      pkwsheet.addRow(Array(String(arr)));
    });

    const buffer = await pkw.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = "PKW.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
  const wptgDownload = async () => {
    const wptg = new ExcelJS.Workbook();
    const wptgsheet = wptg.addWorksheet("Sheet 1");
    const wptgarr = [];

    wptgData.forEach((arr) => {
      wptgarr.push(arr[0], arr[13].match(/\d+/g));
    });

    const filterWptg = [...new Set(wptgarr.flat(Infinity))];
    const filterWptgarr = filterWptg.filter((element) => element !== null);

    filterWptgarr.forEach((arr) => {
      wptgsheet.addRow(Array(String(arr)));
    });

    const buffer = await wptg.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = "WPTG.xlsx";
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
      <div className="flex mt-4 gap-2">
        <button
          onClick={blDownload}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載BL
        </button>
        <button
          onClick={wpkDownload}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載WPK
        </button>
        <button
          onClick={pkwDownload}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載PKW
        </button>
        <button
          onClick={wptgDownload}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          下載WPTG
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
