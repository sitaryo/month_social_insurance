import React, {useState} from 'react';
import PersonStatistics from "./pages/PersonStatistics";
import CompanyStatistics from "./pages/CompanyStatistics";


const App: React.FC = () => {
  const [type, setType] = useState(0);

  return (
    <div style={{display: "flex", padding: 8, flexDirection: "column"}}>
      <h4 style={{textAlign: "center"}}>参保汇总工具 v1.1.0</h4>
      <button onClick={() => setType(type === 0 ? 1 : 0)}>
        {type === 0 ? "方式1" : "方式2"}
      </button>
      {
        type === 0 ?
          <PersonStatistics/> :
          <CompanyStatistics/>
      }
    </div>
  );
}

export default App;
