
const WaitingModal = (props:  { message: string }) => {

    const { message } = props
    return (
        <>
            <dialog id="my-waiting-modal" className="modal">
                <div className="modal-box">
                    <h3 className="font-bold text-lg">Please wait</h3><span className="loading loading-dots loading-md"></span>
                    <p className="py-4">{message}</p>
                </div>
            </dialog>
        </>
    )
}

export default WaitingModal