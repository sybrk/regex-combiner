import { useState } from "react"

const useWaitingHandler = () => {
    const [waitingMessage, setWaitingMessage] = useState("")
    
    function closeModal() {
        (document.getElementById("my-waiting-modal") as HTMLDialogElement)?.close();
    }
    function openModal() {
        (document.getElementById("my-waiting-modal") as HTMLDialogElement)?.showModal();
    }

    return [waitingMessage, setWaitingMessage, openModal, closeModal] as const
}

export default useWaitingHandler