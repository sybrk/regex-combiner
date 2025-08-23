import { useState, useEffect, useRef } from 'react';
import excelFile from "../assets/Regex.Combiner.1.8.xlsm.zip";
import Footer from './footer/Footer';
import { NavLink } from 'react-router';

const HeroCustomized = () => {
  const [isVisible, setIsVisible] = useState({'main-heading': false});

  const observerRef = useRef();
  const [activeFeature, setActiveFeature] = useState(0);

  const features = [
    {
      icon: 'fas fa-file-import',
      title: 'Smart Import',
      description: 'Drag & drop multiple Trados regex files',
      badge: 'Bulk Processing',
      color: 'primary',
      delay: 'delay-200'
    },
    {
      icon: 'fas fa-code',
      title: 'Advanced Editor',
      description: 'Live pattern testing fully compatible with .NET regex flavour.',
      badge: 'Smart Assist',
      color: 'secondary',
      delay: 'delay-300'
    },
    {
      icon: 'fas fa-shield-alt',
      title: 'Validation',
      description: 'Validate your entries before combining and avoid errors in Trados.',
      badge: 'Error-Free',
      color: 'accent',
      delay: 'delay-400'
    },
    {
      icon: 'fas fa-compress-alt',
      title: 'Combine',
      description: 'Create a regex file fully compliant with Trados.',
      badge: 'Smooth operation',
      color: 'info',
      delay: 'delay-500'
    }
  ];

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
    { threshold: 0, rootMargin: '0px 0px -50px 0px' }
  );

  return () => observerRef.current?.disconnect();
}, []);


  // Auto-rotate feature highlight
  useEffect(() => {
    const interval = setInterval(() => {
      setActiveFeature(prev => (prev + 1) % features.length);
    }, 3000);
    return () => clearInterval(interval);
  }, [features.length]);





 






  const observeElement = (element) => {
    if (element && observerRef.current) {
      observerRef.current.observe(element);
    }
  };



  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-purple-900 to-slate-900">

      {/* Hero Section */}
      <div className="hero min-h-screen relative overflow-hidden pt-16">
        {/* Animated background elements */}
        <div className="absolute inset-0 opacity-20">
          <div className="absolute top-20 left-10 w-32 h-32 rounded-full bg-primary/30 animate-pulse"></div>
          <div className="absolute top-40 right-20 w-24 h-24 rounded-full bg-secondary/30 animate-bounce"></div>
          <div className="absolute bottom-20 left-1/4 w-20 h-20 rounded-full bg-accent/30 animate-ping"></div>
        </div>

        <div className="hero-content text-center text-white relative z-10">
          <div className="max-w-6xl">


            {/* Main heading */}
            <div
              id="main-heading"
              ref={observeElement}
              className={`transform transition-all ${
                isVisible['main-heading'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-12'
              }`}
            >
              <h1 className="text-7xl font-black mb-6">
                <span className="bg-gradient-to-r from-primary via-secondary to-accent bg-clip-text text-transparent">
                  Regex Combiner
                </span>
                <br />
                <span className="text-white/90 text-5xl font-light">Made Simple</span>
              </h1>
            </div>

            <p
              id="description"
              ref={observeElement}
              className={"text-xl mb-12 text-white/70 max-w-3xl mx-auto leading-relaxed transform transition-all " + (isVisible['description'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              )}
            >
              The most powerful tool for managing Trados Studio regex files. Import multiple files, edit entries,
              validate regex patterns, and combine into one file seamlessly.
            </p>



            {/* CTA Buttons */}
            <div
              id="cta-buttons"
              ref={observeElement}
              className={`flex flex-col sm:flex-row gap-4 justify-center mb-16 transform transition-all duration-1000 delay-900 ${
                isVisible['cta-buttons'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >

              <NavLink className="btn btn-primary btn-lg gap-2 hover:scale-105 transition-transform hover:shadow-lg hover:shadow-primary/25" to="/combiner" end>
                Try now
              </NavLink>

            </div>
          </div>
        </div>


      </div>

      {/* Features Section */}
      <div className="py-24 bg-base-100">
        <div className="container mx-auto px-6">
          {/* Section Header */}
          <div className="text-center mb-20">

            <h2
              id="features-title"
              ref={observeElement}
              className={`text-5xl font-bold mb-6 transform transition-all duration-1000 delay-200 ${
                isVisible['features-title'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              Everything You Need
            </h2>
            <p
              id="features-desc"
              ref={observeElement}
              className={`text-xl text-base-content/70 max-w-2xl mx-auto transform transition-all duration-1000 delay-400 ${
                isVisible['features-desc'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              Designed specifically for Trados regex file management
            </p>
          </div>

          {/* Feature Cards Grid */}
          <div className="grid auto-cols-auto grid-cols-1 lg:grid-cols-2 gap-8">
            {features.map((feature, index) => (
              <div
                key={feature.title}
                id={`feature-${index}`}
                ref={observeElement}
                className={`card bg-base-200 shadow-xl hover:shadow-2xl transform transition-all duration-700 hover:-translate-y-2 hover:scale-105 ${
                  isVisible[`feature-${index}`] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
                } ${activeFeature === index ? 'ring-2 ring-primary ring-opacity-50' : ''}`}
                style={{ transitionDelay: `${index * 100}ms` }}
              >
                <div className="card-body">
                  <div className="flex items-center mb-4">
                    <div className={`avatar placeholder mr-4 transition-all duration-300 ${activeFeature === index ? 'scale-110' : ''
                      }`}>
                      <div className={`bg-${feature.color} text-${feature.color}-content rounded-xl w-12`}>
                        <i className={`${feature.icon} text-xl`}></i>
                      </div>
                    </div>
                    <h3 className="card-title text-2xl">{feature.title}</h3>
                  </div>
                  <p className="text-base-content/70 mb-6 leading-relaxed">
                    {feature.description}
                  </p>
                  <div className="card-actions">
                    <div className={`badge badge-${feature.color} badge-outline gap-2 hover:scale-105 transition-transform`}>
                      <i className="fas fa-bolt"></i>
                      {feature.badge}
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>



      <div className="py-24 bg-base-100">
        <div className="container mx-auto px-6 text-center">
          <div className="max-w-4xl mx-auto">

            <h2
              id="cta-title"
              ref={observeElement}
              className={`text-5xl font-bold mb-8 transform transition-all duration-1000 delay-200 ${
                isVisible['cta-title'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              Looking for old Excel version?
            </h2>
            <p
              id="cta-desc"
              ref={observeElement}
              className={`text-xl text-base-content/70 mb-12 max-w-2xl mx-auto transform transition-all duration-1000 delay-400 ${
                isVisible['cta-desc'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              While we make the online version better each day, we understand you might still need the Excel version for specific reasons and you can still download it.
            </p>

            <div
              id="final-cta"
              ref={observeElement}
              className={`flex flex-col sm:flex-row gap-4 justify-center mb-8 transform transition-all duration-1000 delay-600 ${
                isVisible['final-cta'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              <NavLink className="btn btn-primary btn-lg gap-2 hover:scale-105 transition-all hover:shadow-lg hover:shadow-primary/25" to="/combiner" end>

                Start online
              </NavLink>
              <a href={excelFile} download={"Regex.Combiner.1.8.xlsm.zip"} className="btn btn-outline btn-lg gap-2 hover:scale-105 transition-all">

                Download excel
              </a>
            </div>

            <div
              id="trust-badges"
              ref={observeElement}
              className={`text-sm text-base-content/50 transform transition-all duration-1000 delay-800 ${
                isVisible['trust-badges'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'
              }`}
            >
              Excel version might not be fully compliant with .NET regex flavour.
            </div>
          </div>
        </div>
      </div>

      <Footer />
    </div>
  );
};

export default HeroCustomized;