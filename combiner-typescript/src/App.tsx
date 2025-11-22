import { Route, Routes } from 'react-router'
import './App.css'
import Navbar from './components/navigation/Navbar'
import Combiner from './components/combiner/Combiner'
import HeroCustomized from './components/HeroCustomized'
import Support from './components/support/Support'

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
