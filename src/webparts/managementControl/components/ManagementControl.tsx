import * as React from 'react';

import styles from './ManagementControl.module.scss';
import "./ManagementControl.module.scss"

import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import { DataGrid, GridColDef, GridColumnGroupingModel } from '@mui/x-data-grid';
import Typography from '@mui/material/Typography';
import Select, { SelectChangeEvent } from '@mui/material/Select';

import type { IManagementControlProps } from './IManagementControlProps';
import getSP from "../PnPjsConfig";
import { v4 as uuidv4 } from 'uuid';
import { PagedItemCollection } from '@pnp/sp/items';
export interface IManagementControlState {
  products: any[];
  
  itemData: {
    productId: string;
    DateOfPurification: Date,
    SerumNumber: string;
    ICANumber: string;
    ColumnNumber: string;
    TotalQuantity: number;
    ExtraYieldCV: string;
    ColumnPreparationDate: Date;
    LabellingDate: Date;
    BlockingPeptidePreparationDate: Date;
    PeptideSupplier: string;
    IFC: string;
    IHC: string;
    LotNumber: string;
    EmployeeName: string;
  }
}

export default class ManagementControl extends React.Component<IManagementControlProps, IManagementControlState> {
  sp = getSP(this.props.context);

  constructor(props: IManagementControlProps) {
    super(props);

    this.state = {
      products: [],
      itemData: {
        productId: "", 
        DateOfPurification: null as any,
        SerumNumber: "",
        ICANumber: "",
        ColumnNumber: "",
        TotalQuantity: 0,
        ExtraYieldCV: "",
        ColumnPreparationDate: null as any,
        LabellingDate: null as any,
        BlockingPeptidePreparationDate: null as any,
        PeptideSupplier: "",
        IFC: "",
        IHC: "",
        LotNumber: "",
        EmployeeName: "",
      }
    };
  }

  componentDidMount(): void {
    this.onInit();
  }

  async onInit() {
    const products = await this.sp.web.lists.getById(this.props.ProductsListId).items();
    console.log("componentDidMount - products:", products);
    this.setState({ products: products });
  }

  // Generic onChange
  onChange = (e: SelectChangeEvent) => {
    const { name, value } = e.target;  
    
    this.setState({
      itemData: {
        ...this.state.itemData,
        [name]: value
      },
    }, ()=>{
      
      this.fetchProductDataFromLists()
    });
  };

  async fetchProductDataFromLists() {

    const lists = await this.sp.web.lists()
    //console.log("fetchProductDataFromLists - lists:", lists)
    const newItems: any[] = [];

    const getItemsReq = lists.map(async (l: any) => {
      try {
        const items: PagedItemCollection<any[]> = await this.sp.web.lists.getById(l.Id).items.filter(`ProductId eq '${this.state.itemData.productId}'`).getPaged();
        
        newItems.push(...items.results)

        return items;
      } catch (err) {
        return null;
      }
      // if (items?.length) {
      //   items.forEach(i => i.listId = l.Id)
      // }
    })

    Promise.all(getItemsReq).then(res => {
      console.log("fetchProductDataFromLists - newItems:", newItems)
    })
  }

  public render(): React.ReactElement<IManagementControlProps> {
    const { Title } = this.props;

    const columns: GridColDef[] = [
      { field: 'id', headerName: 'ProductId', width: 150 },
      { field: 'dateOfPurification', headerName: 'Date of Purification', width: 150 },
      { field: 'serumNumber', headerName: 'Serum Number', width: 130 },
      { field: 'icaNumber', headerName: 'ICA Number', width: 130 },
      { field: 'columnNumber', headerName: 'Column Number', width: 130 },
      { field: 'totalQuantity', headerName: 'Total Quantity (mg)', width: 180 },
      { field: 'extraYield', headerName: 'Extra Yield [C], (V)', width: 170 },
      { field: 'columnPreparationDate', headerName: 'Column Preparation Date', width: 180 },
      { field: 'labellingDate', headerName: 'Labelling Date', width: 150 },
      { field: 'blockingPeptidePreparationDate', headerName: 'Blocking Peptide Preparation Date', width: 250 },
      { field: 'peptideSupplier', headerName: 'Peptide Supplier', width: 150 },
      { field: 'ifc', headerName: 'IFC', width: 200 },
      { field: 'ihc', headerName: 'IHC', width: 200 },
      { field: 'lotNumber', headerName: 'Lot Number', width: 120 },
      { field: 'employeeName', headerName: 'Employee Name', width: 150 },
    ];
    
    const rows = [
      {
        id: this.state.itemData.productId,
        dateOfPurification: '2024-01-01',
        serumNumber: 'S123',
        icaNumber: 'ICA456',
        columnNumber: 'CN789',
        columnPreparationDate: '2024-01-02',
        totalQuantity: 50,
        extraYield: 'Yes',
        labellingDate: '2024-01-03',
        blockingPeptidePreparationDate: '2024-01-04',
        peptideSupplier: 'Supplier A',
        ifc: 'Yes',
        ihc: 'No',
        lotNumber: 'L1234',
        employeeName: 'John Doe'
      },
      // Add more rows as needed
    ];

    const columnGroupingModel: GridColumnGroupingModel = [
      {
        groupId: 'Development',
        description: 'Development',
        children: [
          {
            groupId: 'Antibody Purification',
            description: 'Antibody Purification',
            children: [
              { field: 'dateOfPurification' },
              { field: 'serumNumber' },
              { field: 'icaNumber' },
              { field: 'columnNumber' },
              { field: 'totalQuantity' },
              { field: 'extraYield' },
            ],
          },
          {
            groupId: 'Column preparation / Column preparation for fusion protein',
            description: 'Column preparation / Column preparation for fusion protein',
            children: [
              { field: 'columnPreparationDate' },
            ],
          },
          {
            groupId: 'Fluorophore Labelling',
            description: 'Fluorophore Labelling',
            children: [
              { field: 'labellingDate' },
            ],
          },
          {
            groupId: 'Blocking peptide preparation / fusion blocking peptide preparation',
            description: 'Blocking peptide preparation / fusion blocking peptide preparation',
            children: [
              { field: 'blockingPeptidePreparationDate' },
              { field: 'peptideSupplier' },
            ],
          },
        ],
      },
      {
        groupId: 'Applications',
        description: 'Applications',
        children: [
          {
            groupId: 'Indirect flow cytometry',
            description: 'Indirect flow cytometry',
            children: [
              { field: 'ifc' },
            ]
          },
          {
            groupId: 'Imuunohistochemistry',
            description: 'Imuunohistochemistry',
            children: [
              { field: 'ihc' },
            ]
          },
          { field: 'lotNumber' },
          { field: 'employeeName' },
        ],
      },
    ];
    
    return (
      <section style={{padding: "1em"}}>
        <h1>{Title}</h1>
        <FormControl sx={{ m: 1, minWidth: 120 }} size="small">
          <InputLabel id="demo-select-small-label">Products</InputLabel>
          <Select
            labelId="demo-select-small-label"
            id="demo-select-small"
            value={this.state.itemData.productId}
            label="Product"
            name={"productId"}
            onChange={this.onChange}
          >
            <MenuItem value="">
              <em>None</em>
            </MenuItem>
            {this.state.products.map((item) => (
              <MenuItem key={uuidv4()} value={item.Id}>{item.ProductName.Description}</MenuItem>
            ))}
          </Select>
        </FormControl>
        
          <div className={styles.container} style={{ height: 800, width: '100%', overflow: "auto" }}>
            <div>
              <DataGrid
                experimentalFeatures={{ columnGrouping: true }}
                rows={rows}
                columns={columns}
                checkboxSelection={false}
                disableRowSelectionOnClick
                columnGroupingModel={columnGroupingModel}
              />
            </div>
      
        </div>
      </section>
    );
  }
}
