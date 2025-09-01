import { useDispatch, useSelector } from "react-redux"
import { updateRegex } from "../features/regexesSlice"

const DetailedRegex = () => {

    const regexObj = useSelector((state) => state.regexes.value["1"])
    const dispatch = useDispatch()
    console.log("objd", regexObj)
    const myPattern = /\[([^-\]])+\]/g
    

    const addSpan = (e) => {
        const value = e.target.value
        dispatch(updateRegex({ id: "1", field: "target", data: value }))


    }
    const highlight = (text) => {
        let newValue = ""
        let segments = []
        //let match = myPattern.exec(text) 
        let match;
        
        let counter = 0
        while((match = myPattern.exec(text)) !== null) {
            console.log("buraya gelin mi")
            const start = match.index
            const end = match[0].length
            newValue += text.slice(counter,start)
            newValue += "<div class='tooltip tooltip-primary' data-tip='guu'><span class='bg-accent text-primary-content'>"
            newValue += text.slice(start, start + end)
            newValue += "</span></div>"
            counter = start + end
            if(match.index == myPattern.lastIndex) {
                myPattern.lastIndex++;
            }
        }
        console.log("matchi", match)
        if (newValue.length) {
            console.log("newvalue", newValue)
            return newValue

        } else {
            console.log("no chamge", text)
            return text
        }
    }

    return (
        <>
            <dialog id="detailed-regex-modal" className="modal">
                <div className="modal-box w-11/12 max-w-7xl">
                    <form method="dialog">
                        {/* if there is a button in form, it will close the modal */}
                        <button className="btn btn-sm btn-circle btn-ghost absolute right-2 top-2">✕</button>
                    </form>
                    <div className="flex flex-row gap-2">
                        <div className="text-primary-content w-1/2" contentEditable={true} role="textarea">
                            {regexObj?.source}
                        </div>
                        <div className="relative w-1/2 mt-2 tracking-widest">
                            <div
                                className="absolute inset-0"

                                dangerouslySetInnerHTML={{ __html: highlight(regexObj?.target) }}
                                
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