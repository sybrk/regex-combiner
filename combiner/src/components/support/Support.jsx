import { useEffect } from "react";

const Support = () => {

    useEffect(() => {
        const script = document.createElement("script");
        script.src = "https://tally.so/widgets/embed.js";
        script.async = true;
        document.body.appendChild(script);
    
        return () => {
          document.body.removeChild(script); // cleanup when component unmounts
        };
      }, []);
    
    
    return(
        <>
        <title>Regex Combiner Support</title>
        
        <style>{` iframe { position: absolute; top: 0; right: 0; bottom: 0; left: 0; border: 0; } `}</style>
        <iframe data-tally-src="https://tally.so/r/n09Wr0?transparentBackground=1" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" title="Regex Combiner Support"></iframe>
        </>
    )
}

export default Support