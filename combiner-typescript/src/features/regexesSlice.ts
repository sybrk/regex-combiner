import { createSelector, createSlice, type PayloadAction } from '@reduxjs/toolkit'

import { saveAs } from 'file-saver'
import { newRegexFile, regexNodeBuilder, type RegexRecord, type RegexRecordCollection } from '../utils/Util'

// Define a type for the slice state
export interface RegexesState {
  ids: string[],
  value: RegexRecordCollection
  duplicates: string[],
  invalidIds: string[],
  page: number,
  pageSize: number
}



// Define the initial state using that type
const initialState: RegexesState = {
  ids: [],
  value: {},
  duplicates: [],
  invalidIds: [],
  page: 1,
  pageSize: 100
}
type RegexPayload<K extends keyof RegexRecord = keyof RegexRecord> = {
  id: string;
  field: K;
  data: RegexRecord[K];
}

export const regexesSlice = createSlice({
  name: 'regexes',
  initialState,
  reducers: {

    updateRegex: (state, action: PayloadAction<RegexPayload>) => {
      //console.log("action", action.payload)
      if(state.value) {
      
        Object.assign(state.value[action.payload.id], {[action.payload.field]: action.payload.data})
      }
      
    },
    importRegexes: (state, action: PayloadAction<RegexRecordCollection>) => {
      //console.log(action.payload)
      if(state.ids.length) {
        const lastId = state.ids[state.ids.length -1];
        const comingIds = Object.keys(action.payload).sort((a,b) => Number(a)-Number(b))
        const newIds = comingIds.map(x => (parseInt(x) + parseInt(lastId)).toString())
        const tmpObj : RegexRecordCollection = {}
        for (let i = 0; i < newIds.length; i++) {
          const element = newIds[i];
          tmpObj[element] = action.payload[comingIds[i]]
        }
        //console.log("ekleniyor", {...tmpObj})
        state.value = {...state.value, ...tmpObj}
        state.ids = Object.keys(state.value).sort((a,b) => Number(a)-Number(b))
      } else {
        state.value = action.payload
        state.ids = Object.keys(action.payload).sort((a,b) => Number(a)-Number(b))
        state.page = 1
      }
      
    },
    newRegex: (state) => {
      //console.log("max", Math.max(Object.keys(state.value)), Object.keys(state.value))
      const newId = (state.value && state.ids.length) ? Math.max(...Object.keys(state.value).map(x => parseInt(x))) + 1 : 0;
      if (!state.value) {
        state.value = {}
      }
       
      

      state.value[String(newId)] = {
        "file": "New Regex",
        "description": "",
        "ignoreCase": "true",
        "source": "",
        "sourceValid": null,
        "target": "",
        "targetValid": null,
        "condition": "TargetAndSource",
        hasIssues: null
      }
      state.ids = [...state.ids, String(newId)]
    },
    removeRegex: (state, action: PayloadAction<string>) => {
      if(state.value) {
        delete state.value[action.payload]
        state.ids = state.ids.filter(x => x != action.payload)
      }
      
    },
    combineRegexes: (state) => {

      if(!state.ids.length || !state.value ) return

      let combine = true;
      
      const allDescriptions = Object.keys(state.value).map(x =>
        state.value ? state.value[x].description : ""
      );
      state.duplicates = allDescriptions.filter((x, i) => x != null && allDescriptions.indexOf(x) !== i)
      
      //reset and start invalidIds
      state.invalidIds = Object.keys(state.value).filter((x,i) => {
        if(state.duplicates.includes(state.value[x].description) || !state.value[x].description.length) {
          return x
        }
        
      })
      //look for invalid regexes
      const invalidRegexes = Object.keys(state.value).filter(x => (state.value[x].sourceValid == false || state.value[x].targetValid == false))
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
        state.ids = [...state.invalidIds]
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
      regexCountNode.textContent = regexCount.toString();
      parent?.appendChild(regexCountNode);

      let regexId = 0;
      state.ids.sort((a,b) => Number(a)-Number(b)).map((regex) => {
        const settingNode = regexNodeBuilder(state.value[regex], regexId.toString());
        parent?.appendChild(settingNode);
        regexId++;
      });

      const serializer = new XMLSerializer();
      const serializedFile = serializer.serializeToString(newFile);
      const fileToDownload = new File([serializedFile], "combined.sdlqasettings", {
        type: "text/xml",
      });
      saveAs(fileToDownload);
    },
    setPage(state, action: PayloadAction<number>) {
      state.page = action.payload;
    },
    resetInvalidIds(state) {
      state.invalidIds = []
    },
    addToInvalidIds(state, action: PayloadAction<string>) {
      if(!state.invalidIds.includes(action.payload)){
        state.invalidIds = [...state.invalidIds, action.payload]
      }
    },
    searchByDescription(state, action: PayloadAction<string>) {
      const result = Object.keys(state.value).filter(x => state.value[x]["description"].search(new RegExp(action.payload, "i")) != -1)
      state.ids = [...result]
      state.page = 1;
    },
    removeFilters(state) {
      
      state.ids = [...Object.keys(state.value)]
      state.page = 1;
    },
    removeAllRegexes() : RegexesState {
      
      return regexesSlice.getInitialState()
    }
  }
})

// Action creators are generated for each case reducer function
export const { updateRegex, importRegexes, newRegex, removeRegex, combineRegexes, setPage, searchByDescription, removeFilters, removeAllRegexes } = regexesSlice.actions

export const selectRegexByIdAndField = (state: RegexesState, regexId: string, field: keyof RegexRecord)  => state.value[regexId][field]
export const selectIds = (state: RegexesState) => state.ids
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

export const selectTotalPages = createSelector(
  state => state.regexes.value,
  state => state.regexes.pageSize,
  (value, pageSize) => {

    return Math.ceil(Object.keys(value).length / pageSize)
  }
);

export default regexesSlice.reducer