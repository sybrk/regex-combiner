const OldExcel = ({isVisible, observeElement}) => {

    return(
        <>
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
              <button className="btn btn-primary btn-lg gap-2 hover:scale-105 transition-all hover:shadow-lg hover:shadow-primary/25">
                
                Start online
              </button>
              <button className="btn btn-outline btn-lg gap-2 hover:scale-105 transition-all">
                
                Download excel
              </button>
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
        </>
    )
}

export default OldExcel