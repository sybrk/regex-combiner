import { configureStore } from '@reduxjs/toolkit'

import regexesReducer from '../features/regexesSlice'
import messagesReducer from '../features/messagesSlice'

export const store =  configureStore({
  reducer: {
    regexes: regexesReducer,
    messages: messagesReducer
  }
})

// Infer the `RootState`,  `AppDispatch`, and `AppStore` types from the store itself
export type RootState = ReturnType<typeof store.getState>
// Inferred type: {posts: PostsState, comments: CommentsState, users: UsersState}
export type AppDispatch = typeof store.dispatch
export type AppStore = typeof store