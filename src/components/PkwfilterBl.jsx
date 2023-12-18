import { useState } from "react";
import * as ExcelJS from "exceljs";

const PkwfilterBl = () => {
  const [wpkData, setWpkData] = useState([]);
  const [blData, setBlData] = useState([]);

  const fileChange = async (e, plat) => {
    const selectedFile = e.target.files[0];

    if (selectedFile) {
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
      plat == "pkw"
        ? setWpkData((prevBlData) => [...prevBlData, extractedData])
        : setBlData((prevBlData) => [...prevBlData, extractedData]);
    }
  };

  const downloadFilterPkw = async () => {
    if (wpkData.length != 0 && blData.length != 0) {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      const filterarr = wpkData.flat(Infinity).filter(
        (item) =>
          !blData
            .flat(1)
            .map((subArray) => subArray[0])
            .includes(Number(item))
      );

      const filterPlat = [...new Set(filterarr.flat(Infinity))];
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
      link.download = `filterPkw.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      location.reload();
    }
  };

  return (
    <div className="p-6 mt-20">
      <h2 className="text-2xl font-semibold mb-4">過濾PKW是否跟BL衝突</h2>
      <p>左方放入PKW，右方放入第三方</p>
      <div className="flex gap-4">
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => {
            fileChange(e, "pkw");
          }}
          className="mt-2 p-2 border border-gray-600 rounded-md bg-gray-700 text-gray-200"
        />
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => {
            fileChange(e, "bl");
          }}
          className="mt-2 p-2 border border-gray-600 rounded-md bg-gray-700 text-gray-200"
        />
      </div>

      <button
        onClick={() => {
          downloadFilterPkw();
        }}
        className="bg-green-500 text-white py-2 px-4 mt-4 rounded-md hover:bg-green-600 focus:outline-none focus:ring focus:border-green-300"
      >
        下載過濾後的PKW名單
      </button>
    </div>
  );
};

export default PkwfilterBl;
