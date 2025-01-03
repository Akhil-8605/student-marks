import { BrowserRouter as Router, Route, Routes } from "react-router-dom";
import "./App.css";
import SchoolResultChecker from "./SchoolResultChecker";

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<SchoolResultChecker />} />
        <Route path="*" element={<SchoolResultChecker />} />
      </Routes>
    </Router>
  );
}
export default App;
