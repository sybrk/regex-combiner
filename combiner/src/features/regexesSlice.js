import { createSelector, createSlice } from '@reduxjs/toolkit'

export const regexesSlice = createSlice({
  name: 'regexes',
  initialState: {
    value: {}
  },
  reducers: {
    
    updateRegex: (state, action) => {
      console.log("action", action.payload)
      state.value[action.payload.id][action.payload.field] = action.payload.data
    },
    importRegexes: (state, action) => {
      console.log(action.payload)
        state.value =  action.payload
    }
  }
})

// Action creators are generated for each case reducer function
export const { updateRegex, importRegexes } = regexesSlice.actions
export const getRegexIds = createSelector(
  [state => state.regexes?.value || {}],
  (regexesValue) => Object.keys(regexesValue)
)
export const selectRegexById = (state, regexId) => state.regexes?.value[regexId]
export default regexesSlice.reducer