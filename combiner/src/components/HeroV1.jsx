import { useState, useEffect, useRef } from 'react';

const HeroV1 = () => {

  // this was created by claude ai

  const [isVisible, setIsVisible] = useState({});
  const [activeFeature, setActiveFeature] = useState(0);
  const [stats, setStats] = useState({ files: 0, accuracy: 0, speedup: 0, users: 0 });
  const observerRef = useRef();

  const features = [
    {
      icon: 'fas fa-file-import',
      title: 'Smart Import',
      description: 'Drag & drop multiple Trados regex files with automatic format detection and bulk processing.',
      badge: 'Bulk Processing',
      color: 'primary',
      delay: 'delay-200'
    },
    {
      icon: 'fas fa-code',
      title: 'Advanced Editor',
      description: 'Syntax highlighting, auto-completion, and live pattern testing for professional editing.',
      badge: 'Smart Assist',
      color: 'secondary',
      delay: 'delay-300'
    },
    {
      icon: 'fas fa-shield-alt',
      title: 'Live Validation',
      description: 'Real-time pattern testing and validation to catch errors before they cause problems.',
      badge: 'Error-Free',
      color: 'accent',
      delay: 'delay-400'
    },
    {
      icon: 'fas fa-compress-alt',
      title: 'Intelligent Merge',
      description: 'Smart conflict resolution and duplicate detection for perfect merges every time.',
      badge: 'AI-Powered',
      color: 'info',
      delay: 'delay-500'
    },
    {
      icon: 'fas fa-download',
      title: 'Universal Export',
      description: 'Export to any format, compatible with all major CAT tools and translation software.',
      badge: 'Universal',
      color: 'warning',
      delay: 'delay-600'
    },
    {
      icon: 'fas fa-history',
      title: 'Version Control',
      description: 'Complete change tracking and version history with auto-save functionality.',
      badge: 'Auto-Save',
      color: 'success',
      delay: 'delay-700'
    }
  ];

  const codeLines = [
    { prefix: '$', content: 'npm install regex-merge', type: 'success' },
    { prefix: '>', content: 'Initializing workspace...', type: 'warning' },
    { prefix: '>', content: 'Loading regex patterns...', type: 'info' },
    { prefix: '✓', content: 'Ready to merge!', type: 'success' }
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
      { threshold: 0.1, rootMargin: '0px 0px -50px 0px' }
    );

    return () => observerRef.current?.disconnect();
  }, []);

  // Stats counter animation
  useEffect(() => {
    const timer = setTimeout(() => {
      setStats({
        files: 1000000,
        accuracy: 99.9,
        speedup: 10,
        users: 50000
      });
    }, 1000);
    return () => clearTimeout(timer);
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

  const AnimatedCounter = ({ end, suffix = '', duration = 2000 }) => {
    const [count, setCount] = useState(0);
    
    useEffect(() => {
      if (end === 0) return;
      
      const increment = end / (duration / 50);
      const timer = setInterval(() => {
        setCount(prev => {
          const next = prev + increment;
          return next >= end ? end : next;
        });
      }, 50);
      
      return () => clearInterval(timer);
    }, [end, duration]);
    
    return (
      <span>
        {typeof end === 'number' && end < 100 
          ? count.toFixed(1) 
          : Math.floor(count).toLocaleString()}{suffix}
      </span>
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-purple-900 to-slate-900">
      {/* Navbar */}
      <div className="navbar bg-base-100/10 backdrop-blur-md border-b border-white/10 fixed top-0 z-50">
        <div className="navbar-start">
          <div className="dropdown">
            <div tabindex="0" role="button" className="btn btn-ghost lg:hidden">
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 6h16M4 12h16M4 18h7"></path>
              </svg>
            </div>
            <ul tabindex="0" className="menu menu-sm dropdown-content mt-3 z-[1] p-2 shadow bg-base-100 rounded-box w-52">
              <li><a>Features</a></li>
              <li><a>Pricing</a></li>
              <li><a>Documentation</a></li>
              <li><a>Support</a></li>
            </ul>
          </div>
          <a className="btn btn-ghost text-xl font-bold text-white">
            <div className="avatar placeholder mr-2">
              <div className="bg-primary text-primary-content rounded-lg w-8">
                <span className="text-lg font-bold">R</span>
              </div>
            </div>
            RegexMerge
          </a>
        </div>
        <div className="navbar-center hidden lg:flex">
          <ul className="menu menu-horizontal px-1 text-white">
            <li><a className="hover:text-primary transition-all duration-300">Features</a></li>
            <li><a className="hover:text-primary transition-all duration-300">Pricing</a></li>
            <li><a className="hover:text-primary transition-all duration-300">Docs</a></li>
            <li><a className="hover:text-primary transition-all duration-300">Support</a></li>
          </ul>
        </div>
        <div className="navbar-end">
          <button className="btn btn-ghost btn-circle mr-2 text-white hover:bg-white/10">
            <i className="fab fa-github text-lg"></i>
          </button>
          <button className="btn btn-primary btn-sm hover:scale-105 transition-transform">
            <i className="fas fa-rocket mr-1"></i>
            Get Started
          </button>
        </div>
      </div>

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
            {/* Code mockup */}
            <div 
              id="code-mockup"
              ref={observeElement}
              className={`mockup-code mb-8 max-w-md mx-auto transform transition-all duration-1000 ${
                isVisible['code-mockup'] ? 'opacity-100 translate-y-0 scale-100' : 'opacity-0 translate-y-8 scale-95'
              }`}
            >
              {codeLines.map((line, index) => (
                <pre 
                  key={index} 
                  data-prefix={line.prefix} 
                  className={`text-${line.type} transition-all duration-500`}
                  style={{ transitionDelay: `${index * 200}ms` }}
                >
                  <code>{line.content}</code>
                </pre>
              ))}
            </div>

            {/* Main heading */}
            <div 
              id="main-heading"
              ref={observeElement}
              className={`transform transition-all duration-1000 delay-300 ${
                isVisible['main-heading'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-12'
              }`}
            >
              <h1 className="text-7xl font-black mb-6">
                <span className="bg-gradient-to-r from-primary via-secondary to-accent bg-clip-text text-transparent">
                  Regex Merge
                </span>
                <br />
                <span className="text-white/90 text-5xl font-light">Made Simple</span>
              </h1>
            </div>

            <p 
              id="description"
              ref={observeElement}
              className={`text-xl mb-12 text-white/70 max-w-3xl mx-auto leading-relaxed transform transition-all duration-1000 delay-500 ${
                isVisible['description'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              The most powerful tool for managing Trados regex files. Import multiple files, edit with precision, 
              validate patterns, and combine everything seamlessly. Built for professionals who demand excellence.
            </p>

            {/* Feature badges */}
            <div 
              id="feature-badges"
              ref={observeElement}
              className={`flex flex-wrap justify-center gap-3 mb-12 transform transition-all duration-1000 delay-700 ${
                isVisible['feature-badges'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              {['Multi-Import', 'Smart Editor', 'Validation', 'Auto-Merge'].map((feature, index) => (
                <div 
                  key={feature}
                  className={`badge badge-${['primary', 'secondary', 'accent', 'info'][index]} badge-lg gap-2 
                    hover:scale-110 transition-all duration-300 cursor-pointer`}
                  style={{ transitionDelay: `${index * 100}ms` }}
                >
                  <i className={`fas fa-${['upload', 'edit', 'check-circle', 'layer-group'][index]}`}></i>
                  {feature}
                </div>
              ))}
            </div>

            {/* CTA Buttons */}
            <div 
              id="cta-buttons"
              ref={observeElement}
              className={`flex flex-col sm:flex-row gap-4 justify-center mb-16 transform transition-all duration-1000 delay-900 ${
                isVisible['cta-buttons'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              <button className="btn btn-primary btn-lg gap-2 hover:scale-105 transition-transform hover:shadow-lg hover:shadow-primary/25">
                <i className="fas fa-play"></i>
                Try Demo
              </button>
              <button className="btn btn-outline btn-lg gap-2 text-white border-white hover:bg-white hover:text-gray-900 hover:scale-105 transition-all">
                <i className="fab fa-github"></i>
                View on GitHub
              </button>
            </div>
          </div>
        </div>

        {/* Scroll indicator */}
        <div className="absolute bottom-8 left-1/2 transform -translate-x-1/2 animate-bounce">
          <i className="fas fa-chevron-down text-2xl text-white/50"></i>
        </div>
      </div>

      {/* Features Section */}
      <div className="py-24 bg-base-100">
        <div className="container mx-auto px-6">
          {/* Section Header */}
          <div className="text-center mb-20">
            <div 
              id="features-badge"
              ref={observeElement}
              className={`badge badge-primary badge-lg mb-4 transform transition-all duration-700 ${
                isVisible['features-badge'] ? 'opacity-100 scale-100' : 'opacity-0 scale-95'
              }`}
            >
              <i className="fas fa-star mr-2"></i>
              Features
            </div>
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
              Powerful tools designed specifically for Trados regex file management
            </p>
          </div>

          {/* Feature Cards Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-3 gap-8">
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
                    <div className={`avatar placeholder mr-4 transition-all duration-300 ${
                      activeFeature === index ? 'scale-110' : ''
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

      {/* Stats Section */}
      <div className="py-20 bg-gradient-to-r from-primary to-secondary">
        <div className="container mx-auto px-6">
          <div 
            id="stats"
            ref={observeElement}
            className={`stats stats-vertical lg:stats-horizontal shadow-2xl w-full bg-base-100/10 backdrop-blur-md border border-white/20 transform transition-all duration-1000 ${
              isVisible['stats'] ? 'opacity-100 scale-100' : 'opacity-0 scale-95'
            }`}
          >
            <div className="stat place-items-center text-white">
              <div className="stat-title text-white/70">Files Processed</div>
              <div className="stat-value text-white">
                <AnimatedCounter end={stats.files} suffix="+" />
              </div>
              <div className="stat-desc text-white/60">Daily processing volume</div>
            </div>
            
            <div className="stat place-items-center text-white">
              <div className="stat-title text-white/70">Success Rate</div>
              <div className="stat-value text-white">
                <AnimatedCounter end={stats.accuracy} suffix="%" />
              </div>
              <div className="stat-desc text-white/60">Pattern accuracy</div>
            </div>
            
            <div className="stat place-items-center text-white">
              <div className="stat-title text-white/70">Time Saved</div>
              <div className="stat-value text-white">
                <AnimatedCounter end={stats.speedup} suffix="x" />
              </div>
              <div className="stat-desc text-white/60">Faster than manual</div>
            </div>
            
            <div className="stat place-items-center text-white">
              <div className="stat-title text-white/70">Happy Users</div>
              <div className="stat-value text-white">
                <AnimatedCounter end={stats.users} suffix="+" />
              </div>
              <div className="stat-desc text-white/60">Worldwide professionals</div>
            </div>
          </div>
        </div>
      </div>

      {/* CTA Section */}
      <div className="py-24 bg-base-100">
        <div className="container mx-auto px-6 text-center">
          <div className="max-w-4xl mx-auto">
            <div 
              id="cta-badge"
              ref={observeElement}
              className={`badge badge-primary badge-lg mb-6 transform transition-all duration-700 ${
                isVisible['cta-badge'] ? 'opacity-100 scale-100' : 'opacity-0 scale-95'
              }`}
            >
              <i className="fas fa-rocket mr-2"></i>
              Ready to Launch
            </div>
            <h2 
              id="cta-title"
              ref={observeElement}
              className={`text-5xl font-bold mb-8 transform transition-all duration-1000 delay-200 ${
                isVisible['cta-title'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              Start Merging Like a Pro
            </h2>
            <p 
              id="cta-desc"
              ref={observeElement}
              className={`text-xl text-base-content/70 mb-12 max-w-2xl mx-auto transform transition-all duration-1000 delay-400 ${
                isVisible['cta-desc'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              Join thousands of translation professionals who trust RegexMerge for their daily workflow.
            </p>
            
            <div 
              id="final-cta"
              ref={observeElement}
              className={`flex flex-col sm:flex-row gap-4 justify-center mb-8 transform transition-all duration-1000 delay-600 ${
                isVisible['final-cta'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
              }`}
            >
              <button className="btn btn-primary btn-lg gap-2 hover:scale-105 transition-all hover:shadow-lg hover:shadow-primary/25">
                <i className="fas fa-play"></i>
                Start Free Trial
              </button>
              <button className="btn btn-outline btn-lg gap-2 hover:scale-105 transition-all">
                <i className="fas fa-calendar"></i>
                Book Demo
              </button>
            </div>
            
            <div 
              id="trust-badges"
              ref={observeElement}
              className={`text-sm text-base-content/50 transform transition-all duration-1000 delay-800 ${
                isVisible['trust-badges'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'
              }`}
            >
              <i className="fas fa-shield-alt mr-1"></i>
              No credit card required • 14-day free trial • Cancel anytime
            </div>
          </div>
        </div>
      </div>

      {/* Footer */}
      <footer className="footer footer-center p-10 bg-base-300">
        <div 
          id="footer-content"
          ref={observeElement}
          className={`transform transition-all duration-1000 ${
            isVisible['footer-content'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
          }`}
        >
          <div className="flex items-center gap-2 mb-4">
            <div className="avatar placeholder">
              <div className="bg-primary text-primary-content rounded-lg w-8">
                <span className="text-lg font-bold">R</span>
              </div>
            </div>
            <span className="text-xl font-bold">RegexMerge</span>
          </div>
          <p className="font-medium text-base-content/70">
            Professional Trados Regex File Management
            <br />Built with ❤️ for Translation Professionals
          </p>
          <p className="text-base-content/50 text-sm">© 2025 RegexMerge. All rights reserved.</p>
        </div>
        <div 
          id="footer-social"
          ref={observeElement}
          className={`transform transition-all duration-1000 delay-200 ${
            isVisible['footer-social'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
          }`}
        >
          <div className="grid grid-flow-col gap-4">
            {['twitter', 'github', 'linkedin', 'discord'].map((social) => (
              <a 
                key={social}
                className="link link-hover text-2xl hover:scale-110 hover:text-primary transition-all duration-300"
              >
                <i className={`fab fa-${social}`}></i>
              </a>
            ))}
          </div>
        </div>
      </footer>
    </div>
  );
};

export default HeroV1;