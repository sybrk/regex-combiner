import { NavLink } from "react-router"


const Hero = ({ isVisible, observeElement }) => {


    return (
        <>
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
                            className={"transform transition-all duration-1000 delay-300 " +
                                (isVisible['main-heading']
                                    ? "opacity-100 translate-y-0"
                                    : "opacity-0 translate-y-12")}
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
                            className={"text-xl mb-12 text-white/70 max-w-3xl mx-auto leading-relaxed transform transition-all duration-1000 delay-500 " + (isVisible['description'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
                            )}
                        >
                            The most powerful tool for managing Trados Studio regex files. Import multiple files, edit entries,
                            validate regex patterns, and combine into one file seamlessly.
                        </p>



                        {/* CTA Buttons */}
                        <div
                            id="cta-buttons"
                            ref={observeElement}
                            className={"flex flex-col sm:flex-row gap-4 justify-center mb-16 transform transition-all duration-1000 delay-900 " + (isVisible['cta-buttons'] ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-8'
                            )}
                        >

                            <NavLink className="btn btn-primary btn-lg gap-2 hover:scale-105 transition-transform hover:shadow-lg hover:shadow-primary/25" to="/combiner" end>
                                Try now
                            </NavLink>

                        </div>
                    </div>
                </div>


            </div>
        </>
    )
}

export default Hero