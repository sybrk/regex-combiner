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
      state.ids = Object.keys(action.payload)
    },
    newRegex: (state) => {
      const newId = generateIdFiveChar()
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

      let combine = true;
      const allDescriptions = Object.keys(state.value).map(x =>
        state.value[x]["description"]
      )
      state.duplicates = allDescriptions.filter((x, i) => allDescriptions.indexOf(x) !== i)

      //look for invalid regexes
      const hasInvalidRegex = Object.keys(state.value).some(x => (state.value[x]["sourceValid"] == false || state.value[x]["targetValid"] == false))
      

      // check emptyDescriptions
      if (allDescriptions.some((x) => !x.length)) {
        combine = false
      }

      //check duplicate descriptions
      if (state.duplicates?.length) {
        combine = false
      }

      //check invalid regexes
      if (hasInvalidRegex) {
        combine = false
      }
      if(!combine) {
        window.alert("there are errors to be fixed")
        return
      }

      //create combined file
      const newFile = newRegexFile();
      const parent = newFile.querySelector("SettingsGroup");
      const regexCount = Object.keys(state.value).length;
      const regexCountNode = document.createElementNS("", "Setting");
      regexCountNode.setAttribute("Id", "RegExRulesCount");
      regexCountNode.textContent = regexCount;
      parent.appendChild(regexCountNode);

      let regexId = 0;
      Object.keys(state.value).map((regex) => {
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
    }
  }
})

// Action creators are generated for each case reducer function
export const { updateRegex, importRegexes, newRegex, removeRegex, combineRegexes, setPage } = regexesSlice.actions

export const selectRegexByIdAndField = (state, regexId, field) => state.regexes?.value[regexId][field]
export const selectIds = (state) => state.regexes?.ids
export const selectPagedRegexes = createSelector(
  state => state.regexes.value,
  state => state.regexes.page,
  state => state.regexes.pageSize,
  (value, page, pageSize) => {
    const all = Object.keys(value);
    const start = (page - 1) * pageSize;
    return all.slice(start, start + pageSize);
  }
);
export const selectTotalPages = (state) => {
  const { value, pageSize } = state.regexes;
  return Math.ceil(Object.keys(value).length / pageSize);
};
export default regexesSlice.reducer