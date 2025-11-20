import regexIcon from "../../assets/regexIcon.svg"

const Footer = () => {

    return (
        <>

            <footer className="footer footer-horizontal footer-center bg-primary text-primary-content p-10">
                <aside>
                    <div className="avatar placeholder mr-2">
                        <div className="text-primary-content rounded-lg w-8">
                        <img src={regexIcon} />
                        </div>
                    </div>
                    <p className="font-bold">
                        Regex Combiner
                        <br />
                        The most powerful tool for managing Trados Studio regex files.
                    </p>
                    <p>Copyright © {new Date().getFullYear()} - All right reserved</p>
                </aside>
               
            </footer>
            {/* Footer */}

        </>
    )
}

export default Footer