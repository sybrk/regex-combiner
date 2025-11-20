import { Route, Routes } from 'react-router'
import './App.css'
import Navbar from './components/navigation/Navbar'

function App() {


  return (
    <>
      <Navbar />
      <Routes>
        <Route index element={<HeroCustomized />} />
        <Route path="combiner" element={<Combiner />} />
        <Route path="support" element={<Support />} />
      </Routes>
    </>
  )
}

export default App
