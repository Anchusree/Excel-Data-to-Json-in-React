import { useRef, useState } from 'react'
import './App.css'
import 'bootstrap/dist/css/bootstrap.min.css';
import * as XLSX from 'xlsx'

function App() {

  const [data, setData] = useState([])
  const [fileName, setFileName] = useState('')
  const [searchTerm, setSearchTerm] = useState('')
  const [searchResults, setSearchResults] = useState([])

  const fileRef = useRef()


  //check valid excel file
  const isExcelFile = (file) => {
    const allowedExtensions = ['.xlsx', '.xls']
    const fileName = file.name
    const fileExtension = fileName.slice(fileName.lastIndexOf('.')).toLowerCase()
    return allowedExtensions.includes(fileExtension)

  }

  const handleFile = (e) => {
    const file = e.target.files[0]
    if (!file) return

    if (isExcelFile(file) && file.size != 0) {
      const reader = new FileReader()
      reader.readAsBinaryString(file)
      reader.onload = (e) => {
        const data1 = e.target.result
        const workbook = XLSX.read(data1, { type: "binary" })
        const sheetName = workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]
        const parsedData = XLSX.utils.sheet_to_json(sheet, { defval: null })
        //console.log(parsedData);
        setData(parsedData)
      }
      setFileName(file.name)
    }
    else {
      alert("Invalid file format")
    }

  }

  const handleRemoveFile = () => {
    setFileName(null)
    fileRef.current.value = ""
    setData([])
  }


  const handleSearch = () => {
    const results = []

    data.filter(item => {
      if (item["Name"].toLowerCase().startsWith(searchTerm.toLowerCase()) ||
        item["Name"].toLowerCase().includes(searchTerm.toLowerCase())) {
        results.push(item)

      }
    })
    setSearchResults(results)

  }


  const handleDownload =()=>{

    let keyvalues

    if(data && data.length > 0){
      keyvalues = data.map(item=>item)
    }
    if(searchResults && searchResults.length > 0){
      keyvalues = searchResults.map(item=>item)
    }
    
    const newdata = [...keyvalues]
    
    const worksheet = XLSX.utils.json_to_sheet(newdata)

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook,worksheet,"Sheet1")

    const excelBuffer = XLSX.write(workbook,{ bookType: 'xlsx', type: 'array' })//generates excel
    const excelBlob = new Blob([excelBuffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})

    //download file
    const downloadLink = URL.createObjectURL(excelBlob)
    const link = document.createElement('a')
    link.href= downloadLink
    link.download  = "newdata.xlsx"
    link.click()

    setTimeout(() => {
      URL.revokeObjectURL(downloadLink)
    }, 100);
  }



  return (
    <div>
      <h1>Excel in React</h1>
      <br />

      <div className="mb-3">
        <label htmlFor="formFile" className="form-label">Upload file</label>
        <input className="form-control" type="file" id="formFile" accept=".xlsx,.xls" onChange={handleFile} ref={fileRef} />
        {
          fileName && (
            <>
              <span>{fileName}</span>
              <button className='removeFile' onClick={handleRemoveFile}>X</button>
            </>
          )
        }
      </div>

      <br />


      {
        data && data.length > 0 && (
          <div>
            <h3>Results</h3>
            <div style={{ display: 'flex', justifyContent: 'center' }}>
              <input type='text' placeholder='Search name' value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
              <button className='btn btn-primary' onClick={handleSearch}>Search</button>
            </div>
            <br/>
            <button className='btn btn-success' onClick={handleDownload}>Download</button>

            <br />
            {
              searchResults && searchResults.length > 0
                ?
                <table className="table table-striped">
                  <thead>
                    <tr>
                      {
                        Object.keys(searchResults[0]).map((key) => (
                          <th scope="col" key={key}>{key}</th>
                        ))
                      }

                    </tr>
                  </thead>
                  <tbody>
                    {
                      searchResults.map((row, index) => (
                        <tr key={index}>
                          {Object.values(row).map((value, index) => (
                            <td key={index}>{value}</td>
                          ))}
                        </tr>
                      ))
                    }


                  </tbody>
                </table>
                :
                <table className="table table-striped">
                  <thead>
                    <tr>
                      {
                        Object.keys(data[0]).map((key) => (
                          <th scope="col" key={key}>{key}</th>
                        ))
                      }

                    </tr>
                  </thead>
                  <tbody>
                    {
                      data.map((row, index) => (
                        <tr key={index}>
                          {Object.values(row).map((value, index) => (
                            <td key={index}>{value}</td>
                          ))}
                        </tr>
                      ))
                    }


                  </tbody>
                </table>

            }


          </div>
        )
      }





    </div>
  )
}

export default App
