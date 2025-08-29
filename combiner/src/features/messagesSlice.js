import { createSlice } from '@reduxjs/toolkit'


export const messagesSlice = createSlice({
    name: 'messages',
    initialState: {
        waitingMessage: "",
        showWaitingMessage: false
    },
    reducers: {

        updateWaitingMessage: (state, action) => {

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