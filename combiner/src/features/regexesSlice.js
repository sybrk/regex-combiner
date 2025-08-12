import { createSelector, createSlice } from '@reduxjs/toolkit'
import { generateIdFiveChar } from '../utils/Util'

export const regexesSlice = createSlice({
  name: 'regexes',
  initialState: {
    "ids": [],
    value: {
      
    },
    duplicates : null,
    emptyDescription: null
  },
  reducers: {
    
    updateRegex: (state, action) => {
      console.log("action", action.payload)
      state.value[action.payload.id][action.payload.field] = action.payload.data
    },
    importRegexes: (state, action) => {
      console.log(action.payload)
        state.value =  action.payload
        state.ids = Object.keys(action.payload)
    },
    newRegex: (state) => {
      const newId = generateIdFiveChar()
      state.value[newId] = {
        "file": "New Regex",
        "description": "",
        "ignoreCase": "true",
        "source": "",
        "target": "",
        "condition": "TargetAndSource"
      }
      state.ids = [newId, ...state.ids]
    },
    removeRegex: (state, action) => {
      delete state.value[action.payload.regexId]
      state.ids = state.ids.filter(x => x != action.payload.regexId)
    },
    getDuplicates: (state) => {
      const allDescriptions = Object.keys(state.value).map(x => 
        state.value[x]["description"]
      )
      console.log("all desc", allDescriptions)
      state.duplicates = allDescriptions.filter((x, i) => allDescriptions.indexOf(x) !== i)
    },
    checkEmptyDescriptions: (state) => {
      const allDescriptions = Object.keys(state.value).map(x => 
        state.value[x]["description"]
      )
      
      state.emptyDescription = allDescriptions.some((x, i) => !x.length)
    }
  }
})

// Action creators are generated for each case reducer function
export const { updateRegex, importRegexes, newRegex, removeRegex, getDuplicates, checkEmptyDescriptions } = regexesSlice.actions

export const selectRegexByIdAndField = (state, regexId, field) => state.regexes?.value[regexId][field]
export const selectIds = (state) => state.regexes?.ids
export default regexesSlice.reducer