import { useState } from "react"

const useWaitingHandler = () => {
    const [waitingMessage, setWaitingMessage] = useState("")
    let waitingModal = document.getElementById("my-waiting-modal")
    function closeModal() {
        waitingModal.close();
    }
    function openModal() {
        waitingModal?.showModal();
    }

    return [waitingMessage, setWaitingMessage, openModal, closeModal]
}

export default useWaitingHandler