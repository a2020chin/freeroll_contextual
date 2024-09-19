import { useState } from "react";
import * as ExcelJS from "exceljs";

const ExcelUpload = () => {
  const [file, setFile] = useState([]);
  // const [month, setMonth] = useState(3);
  const [t1filter, setT1filter] = useState(24);
  const [t2filter, setT2filter] = useState(24);
  const [filterDevice, setFilterDevice] = useState("iPhone");
  const [botT1filter, setBotT1filter] = useState(10);
  const [botT2filter, setBotT2filter] = useState(10);
  const [botDeviceIDfilter, setBotDeviceIDfilter] = useState("7.1.2");
  const [botIpfilter, setBotIpfilter] = useState(2);
  const [botDevicefilter, setBotDevicefilter] = useState(2);

  const handleFileChange = async (e) => {
    const selectedFile = e.target.files[0];

    if (selectedFile) {
      // 使用exceljs讀取文件內容
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(selectedFile);
      const worksheet = workbook.worksheets[0];

      // 將數據轉換為陣列

      worksheet.eachRow((row) => {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell) => {
          rowData.push(cell.value);
        });
        setFile((prevData) => [...prevData, rowData]);
      });
    }
  };
  const download = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");
    //const platarr = [];

    //標題
    worksheet.addRow(file[0]);

    file.forEach((arr) => {
      console.log(arr[14] >= t1filter || arr[15] >= t2filter);
      if (
        (arr[14] >= t1filter || arr[15] >= t2filter) &&
        !arr[17]?.includes(filterDevice)
      ) {
        worksheet.addRow(arr);
      }

      if (arr[14] >= botT1filter || arr[15] >= botT2filter) {
        if (String(arr[19])?.includes(botDeviceIDfilter)) {
          if (arr[10] <= botIpfilter && arr[11] <= botDevicefilter)
            worksheet.addRow(arr);
        }
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const blobUrl = URL.createObjectURL(blob);

    // 創建 a 標籤，模擬點擊下載
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = `BotExcelFilter.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className=" p-6">
      <h2 className="text-2xl font-semibold mb-4">BOT config過濾</h2>
      <input
        type="file"
        accept=".xlsx"
        onChange={handleFileChange}
        className="mt-2 p-2 border border-gray-600 rounded-md bg-gray-700 text-gray-200"
      />
      <h3 className="text-lg font-semibold my-4">
        過濾條件一(默認:T1 或 T2 的值大於 24、設備不是使用蘋果系統)
      </h3>
      <div className="flex gap-x-4">
        <div>
          <label htmlFor="t1filter" className="text-white">
            過濾t1：
          </label>
          <input
            type="number"
            id="t1filter"
            value={t1filter}
            onChange={(e) => {
              setBotT1filter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
        <div>
          <label htmlFor="t2filter" className="text-white">
            過濾t2：
          </label>
          <input
            type="number"
            id="t2filter"
            value={t2filter}
            onChange={(e) => {
              setT2filter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
        <div>
          <label htmlFor="filterDevice" className="text-white">
            過濾Device：
          </label>
          <input
            type="text"
            id="filterDevice"
            value={filterDevice}
            onChange={(e) => {
              setFilterDevice(e.target.value);
            }}
            className="bg-gray-700 text-white w-20 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
      </div>

      <h3 className="text-lg font-semibold my-4">
        過濾條件二(默認:T1 或 T2 的值大於 10、設備使用的是 7.1.2 系統、IP
        使用數量小於或等於 2、設備數量小於或等於 2)
      </h3>
      <div className="flex gap-x-4">
        <div>
          <label htmlFor="botT1filter" className="text-white">
            過濾t1：
          </label>
          <input
            type="number"
            id="botT1filter"
            value={botT1filter}
            onChange={(e) => {
              setT1filter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
        <div>
          <label htmlFor="botT2filter" className="text-white">
            過濾t2：
          </label>
          <input
            type="number"
            id="botT2filter"
            value={botT2filter}
            onChange={(e) => {
              setBotT2filter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
        <div>
          <label htmlFor="botDeviceIDfilter" className="text-white">
            過濾設備系統：
          </label>
          <input
            type="text"
            id="botDeviceIDfilter"
            value={botDeviceIDfilter}
            onChange={(e) => {
              setBotDeviceIDfilter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
        <div>
          <label htmlFor="botIpfilter" className="text-white">
            過濾IP數量：
          </label>
          <input
            type="number"
            id="botIpfilter"
            value={botIpfilter}
            onChange={(e) => {
              setBotIpfilter(e.target.value);
            }}
            className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
          />
        </div>
      </div>
      <div>
        <label htmlFor="botDevicefilter" className="text-white">
          過濾Device數量：
        </label>
        <input
          type="number"
          id="botDevicefilter"
          value={botDevicefilter}
          onChange={(e) => {
            setBotDevicefilter(e.target.value);
          }}
          className="bg-gray-700 text-white w-16 p-2 mt-2 border rounded focus:outline-none focus:ring focus:border-blue-300"
        />
      </div>
      <div className="flex mt-4 gap-2">
        <button
          onClick={() => download()}
          className="bg-green-500 text-white py-2 px-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
        >
          過濾下載
        </button>
      </div>
    </div>
  );
};

export default ExcelUpload;
