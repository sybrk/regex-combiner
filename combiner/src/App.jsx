

import { Route, Routes } from 'react-router'
import './App.css'
import Combiner from './Combiner'
import HeroCustomized from './components/HeroCustomized'
import TradosRegexHero from './components/HeroV1'
import Navbar from './components/navigation/Navbar'
import Features from './components/hero/features/Features'
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
