import { useDispatch, useSelector } from "react-redux"
import { updateRegex } from "../features/regexesSlice"

const DetailedRegex = () => {

    const regexObj = useSelector((state) => state.regexes.value["1"])
    const dispatch = useDispatch()
    console.log("objd", regexObj)
    const regexPatterns = [
        {
            regex: /[^.$^{[(|)*+?\\]+/g,
            className: "",
            explanation: (match) => `Literal match ${match[0]}]`
        },
        {
            regex: /\[[^-\]]+\]/g,
            className: "bg-warning/40 text-neutral-content",
            explanation: (match) => `Character class ${match[0]}]`
        },
        {
            regex: /\[[^-\]]+\-[^-\]]+\]/g,
            className: "bg-error",
            explanation: (match) => `Character range ${match[0]}]`
        },
        {
            regex: /\\w/g,
            className: "bg-error",
            explanation: () => `Matches an alpha-numeric character (a-z, A-Z, 0-9, and underscore).`
        },
    ]


    const addSpan = (e) => {
        const value = e.target.value
        dispatch(updateRegex({ id: "1", field: "target", data: value }))


    }
    const highlight = (text) => {
        let newValue = ""
        //let match = myPattern.exec(text) 
        if(text) {
            let match;
            for (let i = 0; i < regexPatterns.length; i++) {
                const pattern = regexPatterns[i];
                let counter = 0
                
                while ((match = pattern.regex.exec(text)) !== null) {
                    console.log("buraya gelin mi")
                    const start = match.index
                    const end = match[0].length
                    newValue += text.slice(counter, start)
                    newValue += `<span class=${pattern.className}>`
                    newValue += text.slice(start, start + end)
                    newValue += "</span>"
                    counter = start + end
                    if (match.index == pattern.regex.lastIndex) {
                        pattern.regex.lastIndex++;
                    }
                }
            }
        }
        

       //console.log("matchi", match)
        if (newValue.length) {
            console.log("newvalue", newValue)
            return newValue

        } else {
            console.log("no chamge", text)
            return text
        }
    }

    const highlightRegex = (text) => {
        if(!text) {
            return
        }
       
        const segments = [];
        let offset = 0;
    
        // Create a copy to track modifications
        let workingText = text;
        
        regexPatterns.forEach((pattern, patternIndex) => {
          let match;
          const regex = new RegExp(pattern.regex);
          
          while ((match = regex.exec(workingText)) !== null) {
            const start = match.index;
            const end = match.index + match[0].length;
            
            segments.push({
              start: start + offset,
              end: end + offset,
              className: pattern.className,
              explanation: pattern.explanation(match),
              text: match[0],
              patternIndex
            });
            
            // Prevent infinite loop
            if (match.index === regex.lastIndex) {
              regex.lastIndex++;
            }
          }
        });
    
        // Sort segments by start position
        segments.sort((a, b) => a.start - b.start);
    
        // Remove overlapping segments (keep first match)
        const filteredSegments = [];
        segments.forEach(segment => {
          if (!filteredSegments.some(existing => 
            (segment.start >= existing.start && segment.start < existing.end) ||
            (segment.end > existing.start && segment.end <= existing.end)
          )) {
            filteredSegments.push(segment);
          }
        });
    
        // Build highlighted text
        let result = '';
        let lastIndex = 0;
    
        filteredSegments.forEach((segment, index) => {
          // Add text before segment
          result += text.slice(lastIndex, segment.start);
          
          // Add highlighted segment
          result += `<span class="${segment.className}">${segment.text}</span>`;
          
          lastIndex = segment.end;
        });
    
        // Add remaining text
        result += text.slice(lastIndex);
    
        return result;
      };

    return (
        <>
            <dialog id="detailed-regex-modal" className="modal">
                <div className="modal-box w-11/12 max-w-7xl">
                    <form method="dialog">
                        {/* if there is a button in form, it will close the modal */}
                        <button className="btn btn-sm btn-circle btn-ghost absolute right-2 top-2">✕</button>
                    </form>
                    <div className="flex flex-row gap-2 bg-">
                        <div className="text-primary-content w-1/2" contentEditable={true} role="textarea">
                            {regexObj?.source}
                        </div>
                        <div className="relative w-1/2 mt-2 tracking-widest">
                            <div
                                className="absolute inset-0"

                                dangerouslySetInnerHTML={{ __html: highlightRegex(regexObj?.target) }}

                            />
                            <textarea
                                className="bg-transparent relative w-full caret-white"
                                value={regexObj?.target}
                                onChange={addSpan}
                            />
                        </div>

                    </div>

                    <p>Hello</p>
                </div>
            </dialog>
        </>
    )
}


export default DetailedRegex