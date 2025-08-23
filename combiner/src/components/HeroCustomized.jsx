import { useState, useEffect, useRef } from 'react';
import Hero from './hero/Hero';
import Features from './hero/features/Features';
import OldExcel from './hero/OldExcel';
import Footer from './footer/Footer';

const HeroCustomized = () => {
  const [isVisible, setIsVisible] = useState({});
  
  const observerRef = useRef();

  

 

  // Intersection Observer for animations
  useEffect(() => {
    observerRef.current = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            setIsVisible(prev => ({ ...prev, [entry.target.id]: true }));
          }
        });
      },
      { threshold: 0.1, rootMargin: '0px 0px -50px 0px' }
    );

    return () => observerRef.current?.disconnect();
  }, []);

 
  
  

  

  const observeElement = (element) => {
    if (element && observerRef.current) {
      observerRef.current.observe(element);
    }
  };

  

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-purple-900 to-slate-900">
      {/* Navbar */}
      

      {/* Hero Section */}
      <Hero isVisible={isVisible} observeElement={observeElement} />

      {/* Features Section */}
      <Features isVisible={isVisible} observeElement={observeElement}  />

      
      {/* CTA Section */}
      <OldExcel isVisible={isVisible} observeElement={observeElement} />

      <Footer />
    </div>
  );
};

export default HeroCustomized;