import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import './App.css'
import { newRegexFile, regexNodeBuilder, regexParser } from './utils/Util'
import { saveAs } from 'file-saver'
import {
  createColumnHelper,
  flexRender,
  getCoreRowModel,
  getSortedRowModel,
  useReactTable,
} from '@tanstack/react-table'
import EditableTextarea from './components/EditableTextarea'
import EditableSelect from './components/EditableSelect'
import TableBody from './components/TableBody'
import { useVirtualizer } from '@tanstack/react-virtual'



function App() {


  const [regexList, setRegexList] = useState([])


  const updateData = useCallback((rowIndex, columnId, value) => {

    setRegexList(old =>
      old.map((row, index) => {
        if (index === rowIndex) {
          return {
            ...old[rowIndex],
            [columnId]: value,
          };
        }
        return row;
      })
    );
  }, []);



  const columnHelper = createColumnHelper();
  // Create cell renderers as useCallback hooks
  const createTextareaCell = useCallback((columnId) => (info) => (
    <EditableTextarea
      key={`${info.row.index}-${columnId}`} // Add key for better React reconciliation
      value={info.getValue()}
      rowIndex={info.row.index}
      columnId={columnId}
      updateData={updateData}
    />
  ), [updateData]);

  const createSelectCell = useCallback((columnId, options) => (info) => (
    <EditableSelect
      key={`${info.row.index}-${columnId}`}
      value={info.getValue()}
      rowIndex={info.row.index}
      columnId={columnId}
      updateData={updateData}
      options={options}
    />
  ), [updateData]);

  const columns = useMemo(() => [
    columnHelper.accessor('file', {
      header: 'File',
      cell: info => info.getValue(),
    }),
    columnHelper.accessor('description', {
      header: 'Description',
      cell: createTextareaCell('description'),
    }),
    columnHelper.accessor('ignoreCase', {
      header: 'Ignore Case',
      cell: createSelectCell('ignoreCase', ['true', 'false']),
    }),
    columnHelper.accessor('source', {
      header: 'Source',
      cell: createTextareaCell('source'),
    }),
    columnHelper.accessor('target', {
      header: 'Target',
      cell: createTextareaCell('target'),
    }),
    columnHelper.accessor('condition', {
      header: 'Condition',
      cell: createSelectCell('condition', [
        'TargetAndSource',
        'TargetNotSource',
        'SourceNotTarget',
        'SourceOnly',
        'TargetOnly',
        'DifferentCount',
        'GroupedSourceNotTarget',
        'GroupedTargetAndSource'
      ]),
    }),
  ], [createTextareaCell, createSelectCell]);

  const table = useReactTable({
    data: regexList,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    meta: {
      updateData,
    },
    debugTable: true
  })

  const { rows } = table.getRowModel()

  const parentRef = useRef(null)
  const virtualizer = useVirtualizer({
    count: rows.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => 34,
    overscan: 20,
  })

  const importRegex = async () => {

    const files = document.getElementById("regexfiles").files
    const regexArr = await regexParser(files)
    setRegexList(regexArr)

  }

  const removeRegex = (e) => {
    const data = e.target.dataset;
    const tmpObj = { ...regexList }
    delete tmpObj[data.fileId]["regexes"][data.regexId]
    setRegexList(tmpObj)
  }

  const createNew = () => {
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


  return (
    <>
      <title>Regex Combiner</title>
      <div className='mt-5 flex flex-row justify-center gap-4 items-center'>
        <input className="file-input file-input-primary" type="file" name="regexfiles" id="regexfiles" multiple />
        <button className="btn btn-primary" onClick={importRegex}>Import</button>
        <p>
          {regexList
            ?
            regexList.length
            :
            0
          }
        </p>
        <button className="btn btn-success rounded-2xl" onClick={createNew}>+</button>
        <button className="btn btn-success rounded-2xl" onClick={combine}>Combine</button>
      </div>

      <div ref={parentRef} style={{ height: `${virtualizer.getTotalSize()}px` }} className="">
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
          <tbody>
            {virtualizer.getVirtualItems().map((virtualRow, index) => {
              const row = rows[virtualRow.index]
              console.log("virtual row", row)
              return (
                <tr
                  key={row.id}
                  style={{
                    height: `${virtualRow.size}px`,
                    transform: `translateY(${virtualRow.start - index * virtualRow.size
                      }px)`,
                  }}
                >
                  {row.getVisibleCells().map((cell) => {
                    return (
                      <td key={cell.id}>
                        {flexRender(
                          cell.column.columnDef.cell,
                          cell.getContext(),
                        )}
                      </td>
                    )
                  })}
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
