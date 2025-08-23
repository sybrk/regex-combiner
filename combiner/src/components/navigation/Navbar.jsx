import { NavLink } from "react-router"

const Navbar = () => {


    return (
        <>
            <div className="navbar bg-linear-65 from-primary to-secondary border-b border-white/10 fixed top-0 z-50 text-base-content">
                <div className="navbar-start">
                    <div className="dropdown">
                        <div tabIndex="0" role="button" className="btn btn-ghost lg:hidden">
                            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 6h16M4 12h16M4 18h7"></path>
                            </svg>
                        </div>
                        <ul tabIndex="0" className="menu menu-sm dropdown-content mt-3 z-[1] p-2 shadow bg-base-100 rounded-box w-52">
                            
                        <li><NavLink to="/combiner" end>
                            Combiner
                        </NavLink></li>
                            <li><NavLink to="/support" end>
                            Support
                        </NavLink></li>
                        </ul>
                    </div>
                    <NavLink to={"/"} end className="btn btn-ghost text-xl font-bold text-white">
                        <div className="avatar placeholder mr-2">
                            <div className="bg-primary text-primary-content rounded-lg w-8">
                                <span className="text-lg font-bold">R</span>
                            </div>
                        </div>

                        
                            RegexCombiner
                        
                    </NavLink>
                </div>
                <div className="navbar-center hidden lg:flex">
                    <ul className="menu menu-horizontal px-1 text-white">
                       

                        <li>
                            <NavLink className="hover:text-primary transition-all duration-300" to="/combiner" end>
                                Combiner
                            </NavLink>
                        </li>
                        <li>
                            <NavLink to="/support" end className="hover:text-primary transition-all duration-300">
                                Support
                            </NavLink>
                        </li>
                    </ul>
                </div>
                <div className="navbar-end">



                </div>
            </div>
        </>
    )
}

export default Navbar