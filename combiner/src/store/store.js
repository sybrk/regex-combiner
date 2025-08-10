import { configureStore } from '@reduxjs/toolkit'

import regexesReducer from '../features/regexesSlice'

export default configureStore({
  reducer: {
    regexes: regexesReducer
  }
})