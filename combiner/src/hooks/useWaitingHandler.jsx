import { useState } from "react"

const useWaitingHandler = () => {
    const [waitingMessage, setWaitingMessage] = useState("")
    
    function closeModal() {
        document.getElementById("my-waiting-modal").close();
    }
    function openModal() {
        document.getElementById("my-waiting-modal").showModal();
    }

    return [waitingMessage, setWaitingMessage, openModal, closeModal]
}

export default useWaitingHandler