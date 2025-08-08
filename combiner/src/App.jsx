import { useCallback, useEffect, useMemo, useState } from 'react'

import './App.css'
import { defaultColumns, idGenerator, newRegexFile, readFile, regexNodeBuilder, xmlParser } from './utils/Util'
import { saveAs } from 'file-saver'
import {
  createColumnHelper,
  flexRender,
  getCoreRowModel,
  getSortedRowModel,
  useReactTable,
} from '@tanstack/react-table'



function App() {


  const [regexList, setRegexList] = useState({})

  const [mydata, setMyData] = useState([])
  const columnHelper = createColumnHelper();
  const columns = useMemo(() => [
    columnHelper.accessor('file', {
      header: 'File',
      cell: info => 
        info.getValue()
      ,
    }),
    columnHelper.accessor('description', {
      header: 'Description',
      cell: info => {
        const initialValue = info.getValue()
        const [value, setValue] = useState(initialValue)
        const onBlur = () => {
          table.options.meta?.updateData(info.row.index, info.column.id, value)
        }
        useEffect(() => {
          setValue(initialValue)
        }, [initialValue])
        return (
          <textarea
            className='textarea'
            value={value}
            onChange={e => setValue(e.target.value)}
            onBlur={onBlur} />
        )
      },
    }),
    columnHelper.accessor('ignoreCase', {
      header: 'Ignore Case',
      cell: info => {
        const initialValue = info.getValue()
        const [value, setValue] = useState(initialValue)
        const onBlur = () => {
          table.options.meta?.updateData(info.row.index, info.column.id, value)
        }
        useEffect(() => {
          setValue(initialValue)
        }, [initialValue])
        return (
          <select
            value={value}
            onChange={e => setValue(e.target.value)}
            onBlur={onBlur}
            className="select select-accent">

            <option>true</option>
            <option>false</option>

          </select>
        )
      },
    }),
    columnHelper.accessor('source', {
      header: 'Source',
      cell: info => {
        const initialValue = info.getValue()
        const [value, setValue] = useState(initialValue)
        const onBlur = () => {
          table.options.meta?.updateData(info.row.index, info.column.id, value)
        }
        useEffect(() => {
          setValue(initialValue)
        }, [initialValue])
        return (
          <textarea
            className='textarea'
            value={value}
            onChange={e => setValue(e.target.value)}
            onBlur={onBlur} />
        )
      },
    }),
    columnHelper.accessor('target', {
      header: 'Target',
      cell: info => {
        const initialValue = info.getValue()
        const [value, setValue] = useState(initialValue)
        const onBlur = () => {
          table.options.meta?.updateData(info.row.index, info.column.id, value)
        }
        useEffect(() => {
          setValue(initialValue)
        }, [initialValue])
        return (
          <textarea
            className='textarea'
            value={value}
            onChange={e => setValue(e.target.value)}
            onBlur={onBlur} />
        )
      },
    }),
    columnHelper.accessor('condition', {
      header: 'Condition',
      cell: info => {
        const initialValue = info.getValue()
        const [value, setValue] = useState(initialValue)
        const onBlur = () => {
          table.options.meta?.updateData(info.row.index, info.column.id, value)
        }
        useEffect(() => {
          setValue(initialValue)
        }, [initialValue])
        return (
          <select
            value={value}
            onChange={e => setValue(e.target.value)}
            onBlur={onBlur}
            className="select select-accent">

            <option>TargetAndSource</option>
            <option>TargetNotSource</option>
            <option>SourceNotTarget</option>
            <option>SourceOnly</option>
            <option>TargetOnly</option>
            <option>DifferentCount</option>
            <option>GroupedSourceNotTarget</option>
            <option>GroupedTargetAndSource</option>

          </select>
        )
      },
    }),
  ], [regexList]);
 
  const table = useReactTable({
    data: mydata,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    meta: {
      updateData: (rowIndex, columnId, value) => {

        setMyData(old =>

          old.map((row, index) => {
            if (index === rowIndex) {
              return {
                ...old[rowIndex],
                [columnId]: value,
              }
            }
            return row
          })
        )
      },
    },
    debugTable: true,
  })

  const importRegex = async () => {
    const tmpObj = { ...regexList }
    const arryData = []
    const files = document.getElementById("regexfiles").files
    for (let index = 0; index < files.length; index++) {

      const file = files[index];
      const fileRead = await readFile(file);
      const parseFile = xmlParser(fileRead.fileContent)
      const fileId = idGenerator();
      tmpObj[fileId] = {}
      tmpObj[fileId]["name"] = fileRead.fileName;

      tmpObj[fileId]["original"] = fileRead.fileContent;
      tmpObj[fileId]["regexes"] = {}
      let regexRules = Array.from(parseFile.querySelectorAll("RegExRule"));
      console.log("regexrules", regexRules)
      regexRules = regexRules.filter(x => /RegExRules\d+$/.test(x.parentElement.getAttribute('Id')))

      regexRules.map((x, i) => {
        tmpObj[fileId]["regexes"][i] = {}
        tmpObj[fileId]["regexes"][i]["description"] = x.querySelector("Description").textContent;
        tmpObj[fileId]["regexes"][i]["ignoreCase"] = x.querySelector("IgnoreCase").textContent;
        tmpObj[fileId]["regexes"][i]["source"] = x.querySelector("RegExSource").textContent;
        tmpObj[fileId]["regexes"][i]["target"] = x.querySelector("RegExTarget").textContent;
        tmpObj[fileId]["regexes"][i]["condition"] = x.querySelector("RuleCondition").textContent;


        const tmpArrObj = {};
        tmpArrObj["file"] = fileRead.fileName;;
        tmpArrObj["description"] = x.querySelector("Description").textContent;
        tmpArrObj["ignoreCase"] = x.querySelector("IgnoreCase").textContent;
        tmpArrObj["source"] = x.querySelector("RegExSource").textContent;
        tmpArrObj["target"] = x.querySelector("RegExTarget").textContent;
        tmpArrObj["condition"] = x.querySelector("RuleCondition").textContent;
        arryData.push(tmpArrObj);
      });

    }
    setRegexList(tmpObj)
    setMyData(arryData)
  }
  const handleChanges = (e) => {
    const data = e.target.dataset;
    const tmpObj = { ...regexList }
    tmpObj[data.fileId]["regexes"][data.regexId][data.field] = e.target.value;
    setRegexList(tmpObj)
  }

  const removeRegex = (e) => {
    const data = e.target.dataset;
    const tmpObj = { ...regexList }
    delete tmpObj[data.fileId]["regexes"][data.regexId]
    setRegexList(tmpObj)
  }

  const createNew = () => {
    /* const tmpObj = { ...regexList }
    if (tmpObj["newRegexes"] == undefined) {
      tmpObj["newRegexes"] = {
        "name": "newRegexes",
        regexes: {}
      }
    }

    const newRegexId = Object.keys(tmpObj["newRegexes"]["regexes"]).length
    tmpObj["newRegexes"]["regexes"][newRegexId] = {
      "description": "",
      ignoreCase: "true",
      source: "",
      target: "",
      condition: "TargetAndSource"
    }
    setRegexList(tmpObj) */
    const newObj = {
      file: "new regex",
      description: "",
      ignoreCase: "false",
      source: "",
      target: "",
      condition: "TargetAndSource"
    }
    setMyData(old => [...old, newObj])
  }
  const combine = () => {
    const newFile = newRegexFile();
    const parent = newFile.querySelector("SettingsGroup");
    const regexCount = Object.keys(regexList).map(x => Object.keys(regexList[x]["regexes"]).length).reduce(((a, b) => a + b), 0);
    const regexCountNode = document.createElementNS("", "Setting");
    regexCountNode.setAttribute("Id", "RegExRulesCount");
    regexCountNode.textContent = regexCount;
    parent.appendChild(regexCountNode)
    let regexId = 0;
    Object.keys(regexList).map((file, i) => {
      Object.keys(regexList[file]["regexes"]).map((regex) => {
        const settingNode = regexNodeBuilder(regexList[file]["regexes"][regex], regexId);
        parent.appendChild(settingNode)
        regexId++
      })
    })
    const serialer = new XMLSerializer();
    const serializedFile = serialer.serializeToString(newFile);
    const fileToDownload = new File([serializedFile], "combined.sdlqasettings", {
      type: "text/xml",
    });
    saveAs(fileToDownload)
  }

  const handleItemChange = useCallback((index, index2, newData) => {
    setRegexList(prevItems => {
      const newItems = { ...prevItems };
      newItems[index]["regexes"][index2]["description"] = newData.description
      return newItems;
    });

  }, []);
  return (
    <>
      <title>Regex Combiner</title>
      <div className='mt-5 flex flex-row justify-center gap-4 items-center'>
        <input className="file-input file-input-primary" type="file" name="regexfiles" id="regexfiles" multiple />
        <button className="btn btn-primary" onClick={importRegex}>Import</button>
        <p>
          {Object.keys(regexList)
            ?
            Object.keys(regexList).map(x => Object.keys(regexList[x]["regexes"]).length).reduce(((a, b) => a + b), 0)
            :
            null
          }
        </p>
        <button className="btn btn-success rounded-2xl" onClick={createNew}>+</button>
        <button className="btn btn-success rounded-2xl" onClick={combine}>Combine</button>
      </div>

      <div className="">
        <table className='table'>
          <thead>
            {table.getHeaderGroups().map(headerGroup => (
              <tr key={headerGroup.id}>
                {headerGroup.headers.map(header => (
                  <th
                    key={header.id}
                    className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                    onClick={header.column.getToggleSortingHandler()}
                  >
                    <div className="flex items-center gap-2">
                      {flexRender(
                        header.column.columnDef.header,
                        header.getContext()
                      )}
                      <span className="text-gray-400">
                        {{
                          asc: '↑',
                          desc: '↓',
                        }[header.column.getIsSorted()] ?? '↕'}
                      </span>
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {table.getRowModel().rows.map(row => {
              console.log("myrow", row)
              return (
                <tr key={row.id} className="hover:bg-gray-50">
                  {row.getVisibleCells().map(cell => (
                    <td key={cell.id} className="px-6 py-4 whitespace-nowrap">
                      {flexRender(
                        cell.column.columnDef.cell,
                        cell.getContext()
                      )}
                    </td>
                  ))}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>


    </>
  )
}

export default App
