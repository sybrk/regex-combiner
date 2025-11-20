import { createSelector, createSlice, type PayloadAction } from '@reduxjs/toolkit'
import { generateIdFiveChar, newRegexFile, regexNodeBuilder } from '../utils/Util'
import { saveAs } from 'file-saver'

// Define a type for the slice state
export interface RegexesState {
  ids: string[],
  value: Record<string, RegexStateType>
  duplicates: string[],
  invalidIds: string[],
  page: number,
  pageSize: number
}

type RegexStateType = {
  file: string,
  description: string,
  ignoreCase: "true" | "false",
  source: string | null,
  sourceValid: boolean | null,
  target: string | null,
  targetValid: boolean | null,
  condition: "TargetAndSource" | "TargetNotSource" | "SourceNotTarget" | "SourceOnly" | "TargetOnly" | "DifferentCount" | "GroupedSourceNotTarget" | "GroupedTargetAndSource",
  hasIssues: boolean | null
}



// Define the initial state using that type
const initialState: RegexesState = {
  ids: [],
  value: {},
  duplicates: [],
  invalidIds: [],
  page: 0,
  pageSize: 100
}
type RegexPayload<K extends keyof RegexStateType = keyof RegexStateType> = {
  id: string;
  field: K;
  data: RegexStateType[K];
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
    importRegexes: (state, action: PayloadAction<Record<string, RegexStateType>>) => {
      //console.log(action.payload)
      if(state.ids.length) {
        const lastId = state.ids[state.ids.length -1];
        const comingIds = Object.keys(action.payload).sort((a,b) => Number(a)-Number(b))
        const newIds = comingIds.map(x => (parseInt(x) + parseInt(lastId)).toString())
        const tmpObj : Record<string, RegexStateType> = {}
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
      parent.appendChild(regexCountNode);

      let regexId = 0;
      state.ids.sort((a,b) => Number(a)-Number(b)).map((regex) => {
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
    },
    searchByDescription(state, action) {
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

export const selectRegexByIdAndField = (state: RegexesState, regexId: string, field: keyof RegexStateType) => state.value[regexId][field]
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
export const selectTotalPages = (state: RegexesState) => {
  const { value, pageSize } = state;
  return Math.ceil(Object.keys(value).length / pageSize);
};
export const selectHasInvalidEntries = (state: RegexesState) => {
  
  return state.invalidIds.length > 0;
};
export default regexesSlice.reducer