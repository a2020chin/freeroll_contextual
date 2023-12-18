// import { useState } from "react";

import ExcelUpload from "./components/ExcelUpload";
import PkwfilterBl from "./components/PkwfilterBl";

function App() {
  return (
    <>
      <div className="w-screen h-screen bg-gray-800 text-white">
        <ExcelUpload />
        <PkwfilterBl />
      </div>
    </>
  );
}

export default App;
