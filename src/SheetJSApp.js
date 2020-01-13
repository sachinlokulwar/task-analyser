/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* Notes:
   - usage: `ReactDOM.render( <SheetJSApp />, document.getElementById('app') );`
   - xlsx.full.min.js is loaded in the head of the HTML page
   - this script should be referenced with type="text/babel"
   - babel.js in-browser transpiler should be loaded before this script
*/

import React from 'react';
import {Table, Container, Label, Input, Button, Dropdown, DropdownToggle, DropdownMenu, DropdownItem} from 'reactstrap';
import XLSX from 'xlsx';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faAngleUp, faAngleDown } from '@fortawesome/free-solid-svg-icons'

class SheetJSApp extends React.Component {
	constructor(props) {
		super(props);
		this.state = {
			data: [], /* Array of Arrays e.g. [["a","b"],[1,2]] */
			cols: [],  /* Array of column objects e.g. { name: "C", K: 2 } */
			usersList: []
		};
		this.handleFile = this.handleFile.bind(this);
		this.processData = this.processData.bind(this);
		this.handleMultipleFiles = this.handleMultipleFiles.bind(this);
	};

	processData(data) {
		if(data.length < 8) {
			return []
		}
		const keys = data[7];
		const response = [];
		for(let i=8; i< data.length;i++) {
			const obj = {}
			for(let j=0;j<keys.length; j++) {
				obj[keys[j]] = data[i][j];
			}
			response.push(obj)
		}
		return response;
	}

	handleFile(file/*:File*/) {
		/* Boilerplate to set up FileReader */
		const reader = new FileReader();
		const rABS = !!reader.readAsBinaryString;
		reader.onload = (e) => {
			/* Parse data */
			const bstr = e.target.result;
			const wb = XLSX.read(bstr, {type:rABS ? 'binary' : 'array'});
			/* Get first worksheet */
			console.log(wb);
			const index = wb.SheetNames.indexOf("Project-Time Registration");
			const wsname = wb.SheetNames[index];
			const ws = wb.Sheets[wsname];
			/* Convert array of arrays */
			const data = XLSX.utils.sheet_to_json(ws, {header:1});
			console.log(data);
			/* Update state */
			const obj = {
				userName: data[2][6],
				data: this.processData(data)
			}
			this.setState({
				data: [...this.state.data, obj],
				usersList: [...this.state.usersList, obj.userName]
			});
		};
		if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file);
	};

	handleMultipleFiles(files) {
		for(let i=0; i< files.length; i++) {
			this.handleFile(files[i]);
		}
	}

	render() { 
		return (
			<Container>
				<DragDropFile handleFile={this.handleFile}>
					<div className="row"><div className="col-xs-12">
						<DataInput handleFile={this.handleMultipleFiles} />
					</div></div>
					<div className="row"><div className="col-12">
						<OutTable data={this.state.data} cols={this.state.cols} usersList={this.state.usersList}/>
					</div></div>
				</DragDropFile>
			</Container>
		);
	};
};

//if(typeof module !== 'undefined') module.exports = SheetJSApp

/* -------------------------------------------------------------------------- */

/*
  Simple HTML5 file drag-and-drop wrapper
  usage: <DragDropFile handleFile={handleFile}>...</DragDropFile>
    handleFile(file:File):void;
*/
export default SheetJSApp;

class DragDropFile extends React.Component {
	constructor(props) {
		super(props);
		this.onDrop = this.onDrop.bind(this);
	};
	suppress(evt) { evt.stopPropagation(); evt.preventDefault(); };
	onDrop(evt) { evt.stopPropagation(); evt.preventDefault();
		const files = evt.dataTransfer.files;
		files && this.props.handleFile(files);
		console.log(files);
	};
	render() { 
		return (
			<div onDrop={this.onDrop} onDragEnter={this.suppress} onDragOver={this.suppress}>
				{this.props.children}
			</div>
		);
	};
};

/*
  Simple HTML5 file input wrapper
  usage: <DataInput handleFile={callback} />
*/
class DataInput extends React.Component {
	constructor(props) {
		super(props);
		this.handleChange = this.handleChange.bind(this);
	};
	handleChange(e) {
		const files = e.target.files;
		files && this.props.handleFile(files);
	};
	render() { 
		return (
			<form className="form-inline">
				<div className="form-group">
					<Label for="exampleFile">Select File: </Label>
					<Input
						type="file"
						name="file"
						id="exampleFile"
						accept={SheetJSFT} 
						onChange={this.handleChange}
						multiple
					/>
				</div>
			</form>
	); };
}

/*
  Simple HTML Table
  usage: <OutTable data={data} cols={cols} />
    data:Array<Array<any> >;
    cols:Array<{name:string, key:number|string}>;
*/
class OutTable extends React.Component {
	constructor(props) { 
		super(props); 
		this.state = {
			combinedData : null,
			selectedKey: null,
			showDetails: false,
			dropdownOpen: false,
			keyDropdownOpen: false,
			selectedFilterKey: "Task ID"
		}
		this.filterData = this.filterData.bind(this);
		this.combineData = this.combineData.bind(this);
		this.processData = this.processData.bind(this);
		this.toggleDetails = this.toggleDetails.bind(this);
		this.showDetailsView = this.showDetailsView.bind(this);
		this.toggle = this.toggle.bind(this);
		this.exportData = this.exportData.bind(this);
		this.togglekeyDropdown = this.togglekeyDropdown.bind(this);
	};

	filterData(data, filterKey) {
		const ids = []
		const response = []
		if(!(data && data.length)) {
			return response;
		}
		for(let i=0; i< data.length; i++) {
			const obj = data[i];
			if(obj[filterKey] && ids.indexOf(obj[filterKey]) === -1) {
				ids.push(obj[filterKey])
			}
		}
		for(let i=0;i<ids.length; i++) {
			const responseObj = {
				key: ids[i],
				value: 0
			}
			for(let j=0; j< data.length; j++) {
				const obj = data[j];
				if(obj[filterKey] === ids[i]) {
					responseObj.value += obj["Sh"];
				}
			}

			response.push(responseObj)
		}
		return response;
	}

	combineData(data) {
		const response = {}
		for(let i=0; i<data.length; i++) {
			const d1 = data[i].filteredData;
			for(let j = 0; j<d1.length; j++) {
				response[d1[j].key] = response[d1[j].key] ? (response[d1[j].key] + d1[j].value) : d1[j].value
			}
		}
		return response;
	}

	processData(data, filterKey) {
		for(let i=0; i<data.length; i++) {
			data[i].filteredData = this.filterData(data[i].data, filterKey)
		}
		this.setState({
			combinedData: this.combineData(data),
			userWiseData: data
		})
	}

	componentWillReceiveProps(nextProps) {
		if(nextProps.data && nextProps.data.length && 
			((this.props.data !== nextProps.data) || (this.props.data && this.props.data.length) !== nextProps.data.length)
		) {
			this.processData(nextProps.data, "Task ID")
		}
	}

	toggleDetails(key) {
		this.setState({
			showDetails: !this.state.showDetails,
			selectedKey: key
		})
	}

	exportData() {
		if(this.state.selectedUser) {
			const selectedUserData = this.state.userWiseData && this.state.userWiseData.find((data) => data.userName === this.state.selectedUser);
			const filData = selectedUserData.filteredData;
			const ws = XLSX.utils.json_to_sheet(filData);
			let new_workbook = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(new_workbook, ws, "SheetJS");
			XLSX.writeFile(new_workbook, this.state.selectedUser + ".xlsx");
		}
		else {
			let new_workbook = XLSX.utils.book_new();
			const keys = Object.keys(this.state.combinedData);
			const list = [];
			for(let i=0; i<keys.length; i++) {
				const obj = {
					key: keys[i],
					value: this.state.combinedData[keys[i]]
				}
				list.push(obj);
			}
			const ws = XLSX.utils.json_to_sheet(list);
			XLSX.utils.book_append_sheet(new_workbook, ws, "total");	
			for(let i = 0; i< this.state.userWiseData.length; i++) {
				const filData = this.state.userWiseData[i].filteredData;
				const ws = XLSX.utils.json_to_sheet(filData);
				XLSX.utils.book_append_sheet(new_workbook, ws, this.state.userWiseData[i].userName);	
			}
			XLSX.writeFile(new_workbook, "combine.xlsx");
		}
		
	}
	showDetailsView(key) {
		let list = [];
		this.state.userWiseData && this.state.userWiseData.filter((data) => {
			const obj = data.filteredData && data.filteredData.find((obj) => obj.key == key)

			if(obj) {
				list.push(<tr><td>{data.userName}</td><td>{parseFloat(obj.value).toFixed(1)}</td></tr>);
			}
			return false
		})
		return list;
	}

	toggle() {
		this.setState({
			dropdownOpen: !this.state.dropdownOpen
		})
	}

	togglekeyDropdown() {
		this.setState({
			keyDropdownOpen: !this.state.keyDropdownOpen
		})
	}

	changeUser(user) {
		this.setState({
			selectedUser: user === 'all' ? null : user
		})
	}

	changeKey(key) {
		this.setState({
			selectedFilterKey : key
		})
		this.processData(this.props.data, key);
	}

	render() {
		//const filteredData = this.filterData(this.props.data && this.props.data[0] && this.props.data[0].data)
		const keys = Object.keys(this.state.combinedData || {});
		console.log(this.props.usersList);
		let selectedUserData = null;
		const processedData = {}
		if(this.state.selectedUser) {
			selectedUserData = this.state.userWiseData && this.state.userWiseData.find((data) => data.userName === this.state.selectedUser);
			const filData = selectedUserData.filteredData;
			for(let i=0;i<filData.length;i++) {
				processedData[filData[i].key] = filData[i].value
			}
		}
		const dataToShow = this.state.selectedUser ? processedData : (this.state.combinedData)
		console.log(dataToShow);
		return (
			<div className="table-responsive" style={{"margin": "20px auto"}}>
				
				<div style={{"margin": "20px", float: "right"}}>
				<div style={{display: "inline"}}>
						<Dropdown isOpen={this.state.keyDropdownOpen} toggle={this.togglekeyDropdown} style={{display: "inline-block"}}>
							<DropdownToggle caret>
								{this.state.selectedFilterKey || "Select Key"}
							</DropdownToggle>
							<DropdownMenu>
								<DropdownItem selected onClick={() => this.changeKey("Task Id")} >Task ID</DropdownItem>
								<DropdownItem onClick={() => this.changeKey("WP ID / RED")} >WP ID / RED</DropdownItem>
							</DropdownMenu>
						</Dropdown>
					</div>
					<div style={{display: "inline", "margin-left": "20px"}}>
						<Dropdown isOpen={this.state.dropdownOpen} toggle={this.toggle} style={{display: "inline-block"}}>
							<DropdownToggle caret>
								{this.state.selectedUser || "Select Name"}
							</DropdownToggle>
							<DropdownMenu>
								<DropdownItem onClick={() => this.changeUser("all")} >All</DropdownItem>
								{
									this.props.usersList && this.props.usersList.map((user) => {
										return(<DropdownItem onClick={() => this.changeUser(user)} >{user}</DropdownItem>)
									})	
								}
							</DropdownMenu>
						</Dropdown>
					</div>
					<div style={{display: "inline", "margin-left": "20px"}}>
						<Button style={{float: "right"}} onClick={this.exportData}>Export</Button>
					</div>
				</div>
				
				<Table bordered style={{"margin-top": "40px auto"}}>
					<thead>
							<tr><th>{this.state.selectedFilterKey}</th><th>Σ H</th></tr>
					</thead>
					<tbody>
						{keys.map((r,i) => <tr key={i}>
							<td>
								{this.state.selectedUser ?
									<>{r}</>
								:
									<Button color="link" onClick={() => this.toggleDetails(r)}>
										{this.state.showDetails && this.state.selectedKey === r ?
											<FontAwesomeIcon icon={faAngleDown}/>
										:
											<FontAwesomeIcon icon={faAngleUp}/>
										}
										&nbsp;{r}
									</Button>
								}
								{this.state.showDetails && this.state.selectedKey === r &&
								<Table bordered>
									<thead>
										<tr><th>Name</th><th>Σ H</th></tr>
									</thead>
									<tbody>
										{this.showDetailsView(this.state.selectedKey)}		
									</tbody>
								</Table>
								}
							</td>
							<td>{parseFloat(dataToShow[r]).toFixed(1)}</td>
						</tr>)}
					</tbody>
				</Table>
			</div>
		);
	};
};

/* list of supported file types */
const SheetJSFT = [
	"xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(function(x) { return "." + x; }).join(",");

/* generate an array of column objects */
const make_cols = refstr => {
	let o = [], C = XLSX.utils.decode_range(refstr).e.c + 1;
	for(var i = 0; i < C; ++i) o[i] = {name:XLSX.utils.encode_col(i), key:i}
	return o;
};