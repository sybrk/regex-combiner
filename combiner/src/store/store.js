import { configureStore } from '@reduxjs/toolkit'

import regexesReducer from '../features/regexesSlice'
import messagesReducer from '../features/messagesSlice'

export default configureStore({
  reducer: {
    regexes: regexesReducer,
    messages: messagesReducer
  }
})