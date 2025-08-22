import { createSelector, createSlice } from '@reduxjs/toolkit'
import { generateIdFiveChar, newRegexFile, regexNodeBuilder } from '../utils/Util'
import { saveAs } from 'file-saver'

export const regexesSlice = createSlice({
  name: 'regexes',
  initialState: {
    "ids": [],
    value: {

    },
    duplicates: null,
    invalidIds : [],
    page: 1,
    pageSize: 100
  },
  reducers: {

    updateRegex: (state, action) => {
      console.log("action", action.payload)
      state.value[action.payload.id][action.payload.field] = action.payload.data
    },
    importRegexes: (state, action) => {
      console.log(action.payload)
      state.value = action.payload
      state.ids = Object.keys(action.payload).sort((a,b) => a-b)
    },
    newRegex: (state) => {
      const newId = Math.max(Object.keys(state.value)) + 1;
      state.value[newId] = {
        "file": "New Regex",
        "description": "",
        "ignoreCase": "true",
        "source": "",
        "sourceValid": null,
        "target": "",
        "targetValid": null,
        "condition": "TargetAndSource"
      }
      state.ids = [...state.ids, newId]
    },
    removeRegex: (state, action) => {
      delete state.value[action.payload.regexId]
      state.ids = state.ids.filter(x => x != action.payload.regexId)
    },
    combineRegexes: (state) => {

      if(!state.ids.length) return

      let combine = true;
      
      const allDescriptions = Object.keys(state.value).map(x =>
        state.value[x]["description"]
      )
      state.duplicates = allDescriptions.filter((x, i) => allDescriptions.indexOf(x) !== i)
      
      //reset and start invalidIds
      state.invalidIds = Object.keys(state.value).filter((x,i) => {
        if(state.duplicates.includes(state.value[x]["description"]) || !state.value[x]["description"].length) {
          return x
        }
        
      })
      //look for invalid regexes
      const invalidRegexes = Object.keys(state.value).filter(x => (state.value[x]["sourceValid"] == false || state.value[x]["targetValid"] == false))
      invalidRegexes.forEach(x =>
        {
          if(!state.invalidIds.includes(x)){
            state.invalidIds = [...state.invalidIds, x]
          }
        }
        
      )

      // check emptyDescriptions
      if (allDescriptions.some((x) => !x.length)) {
        combine = false
      }

      //check duplicate descriptions
      if (state.duplicates?.length) {
        combine = false
      }

      //check invalid regexes
      if (invalidRegexes.length) {
        combine = false
      }
      if(!combine) {
        //console.log("invalidoe", state.invalidIds)
        state.ids = state.invalidIds
        state.page = 1;
        //window.alert("there are errors to be fixed")
       
        return
      }
     
        state.ids = Object.keys(state.value)
        state.page = 1;
      //create combined file
      const newFile = newRegexFile();
      const parent = newFile.querySelector("SettingsGroup");
      const regexCount = Object.keys(state.value).length;
      const regexCountNode = document.createElementNS("", "Setting");
      regexCountNode.setAttribute("Id", "RegExRulesCount");
      regexCountNode.textContent = regexCount;
      parent.appendChild(regexCountNode);

      let regexId = 0;
      state.ids.sort((a,b) => a-b).map((regex) => {
        const settingNode = regexNodeBuilder(state.value[regex], regexId);
        parent.appendChild(settingNode);
        regexId++;
      });

      const serializer = new XMLSerializer();
      const serializedFile = serializer.serializeToString(newFile);
      const fileToDownload = new File([serializedFile], "combined.sdlqasettings", {
        type: "text/xml",
      });
      saveAs(fileToDownload);
    },
    setPage(state, action) {
      state.page = action.payload;
    },
    resetInvalidIds(state) {
      state.invalidIds = []
    },
    addToInvalidIds(state, action) {
      if(!state.invalidIds.includes(action.payload)){
        state.invalidIds = [...state.invalidIds, action.payload]
      }
    }
  }
})

// Action creators are generated for each case reducer function
export const { updateRegex, importRegexes, newRegex, removeRegex, combineRegexes, setPage } = regexesSlice.actions

export const selectRegexByIdAndField = (state, regexId, field) => state.regexes?.value[regexId][field]
export const selectIds = (state) => state.regexes?.ids
export const selectPagedRegexes = createSelector(
  state => state.regexes.ids,
  state => state.regexes.page,
  state => state.regexes.pageSize,
  (ids, page, pageSize) => {
    const all = ids;
    const start = (page - 1) * pageSize;
    return all.slice(start, start + pageSize);
  }
);
export const selectTotalPages = (state) => {
  const { value, pageSize } = state.regexes;
  return Math.ceil(Object.keys(value).length / pageSize);
};
export const selectHasInvalidEntries = (state) => {
  
  return state.regexes?.invalidIds.length > 0;
};
export default regexesSlice.reducer