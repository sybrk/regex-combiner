import { createSlice, type PayloadAction } from '@reduxjs/toolkit'


export interface MessageState {
    waitingMessage: string,
    showWaitingMessage: boolean
}

const initialState: MessageState = {
    waitingMessage : "",
    showWaitingMessage: false
}
export const messagesSlice = createSlice({
    name: 'messages',
    initialState,
    reducers: {

        updateWaitingMessage: (state, action: PayloadAction<string>) => {

            state.waitingMessage = action.payload
        },
        turnOnShowWaitingMessage: (state) => {
            state.showWaitingMessage = true
        },
        turnOffShowWaitingMessage: (state) => {
            state.showWaitingMessage = false
        }
    }
})

// Action creators are generated for each case reducer function
export const { updateWaitingMessage, turnOffShowWaitingMessage, turnOnShowWaitingMessage } = messagesSlice.actions


export default messagesSlice.reducer