import { useEffect, useState } from "react";

const Features = ({isVisible, observeElement}) => {

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

      // Auto-rotate feature highlight
  useEffect(() => {
    const interval = setInterval(() => {
      setActiveFeature(prev => (prev + 1) % features.length);
    }, 3000);
    return () => clearInterval(interval);
  }, [features.length]);

    return (
        <>
            <div className="py-24 bg-base-100">
                <div className="container mx-auto px-6">
                    {/* Section Header */}
                    <div className="text-center mb-20">
                        
                        <h2
                            id="features-title"
                            ref={observeElement}
                            className={`text-5xl font-bold mb-6 transform transition-all duration-1000 delay-200 ${isVisible['features-title'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
                                }`}
                        >
                            Everything You Need
                        </h2>
                        <p
                            id="features-desc"
                            ref={observeElement}
                            className={`text-xl text-base-content/70 max-w-2xl mx-auto transform transition-all duration-1000 delay-400 ${isVisible['features-desc'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
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
                                className={`card bg-base-200 shadow-xl hover:shadow-2xl transform transition-all duration-700 hover:-translate-y-2 hover:scale-105 ${isVisible[`feature-${index}`] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
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
        </>
    )
}

export default Features