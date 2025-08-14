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
    invalidRegexes: false
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
        "target": "",
        "condition": "TargetAndSource"
      }
      state.ids = [newId, ...state.ids]
    },
    removeRegex: (state, action) => {
      delete state.value[action.payload.regexId]
      state.ids = state.ids.filter(x => x != action.payload.regexId)
    },
    combineRegexes: (state) => {


      const allDescriptions = Object.keys(state.value).map(x =>
        state.value[x]["description"]
      )
      state.duplicates = allDescriptions.filter((x, i) => allDescriptions.indexOf(x) !== i)

      //look for invalid regexes
      const allRegexPatterns = [...Object.keys(state.value).map(x =>
        state.value[x]["source"]), ...Object.keys(state.value).map(x =>
          state.value[x]["target"])]
      state.invalidRegexes = false
      for (let index = 0; index < allRegexPatterns.length; index++) {
        const element = allRegexPatterns[index];
        const testString = "Hello World"

        try {
          const regexString = new RegExp(element, "g")
          regexString.test(testString)
        } catch (error) {
          state.invalidRegexes = true
          break
        }
      }

      // check emptyDescriptions
      if (allDescriptions.some((x) => !x.length)) return

      //check duplicate descriptions
      if (state.duplicates?.length) return

      //check invalid regexes
      if (state.invalidRegexes) return

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
    }
  }
})

// Action creators are generated for each case reducer function
export const { updateRegex, importRegexes, newRegex, removeRegex, combineRegexes } = regexesSlice.actions

export const selectRegexByIdAndField = (state, regexId, field) => state.regexes?.value[regexId][field]
export const selectIds = (state) => state.regexes?.ids
export default regexesSlice.reducer