import * as React from 'react';
import { DataGrid, GridToolbar } from '@mui/x-data-grid';
import { v4 as uuidv4 } from 'uuid';
import getSP from "../PnPjsConfig";
import FormControl from '@mui/material/FormControl';
import InputLabel from '@mui/material/InputLabel';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import MenuItem from '@mui/material/MenuItem';
import styles from './ManagementControl.module.scss';
import "./style.css"
import "./ManagementControl.module.scss";
import type { IManagementControlProps } from './IManagementControlProps';
import Skeleton from '@mui/material/Skeleton';

export interface IManagementControlState {
  products: any[];
  productId: string;
  AntibodyPurificationData: {
    DateOfPurification: string,
    SerumNumber: string;
    ICANumber: string;
    ColumnNumber: string;
    TotalQuantity: string;
    ExtraYieldCV: string;
    LotNumber: string;
  };
  AntiBPureRows: any[];
  ColumnPreparationDate: string;
  ColumnPreparationDateRows: any[];
  LabellingDate: string;
  LabellingDateRows: any[];
  peptidePrepData: {
    BlockingPeptidePreparationDate: string,
    PeptideSupplier: string,
  };
  peptidePrepRows: any[];

  IFCLotNumber: string;
  IFCLotNumberRows: any[],
  IHCLotNumber: string;
  IHCLotNumberRows: any[],

  isLoading: boolean;
  skeletonLoading: boolean;
}

export default class ManagementControl extends React.Component<IManagementControlProps, IManagementControlState> {
  sp = getSP(this.props.context);

  constructor(props: IManagementControlProps) {
    super(props);
    this.state = {
      products: [],
      productId: "",
      AntibodyPurificationData: {
        DateOfPurification: "",
        SerumNumber: "",
        ICANumber: "",
        ColumnNumber: "",
        TotalQuantity: "",
        ExtraYieldCV: "",
        LotNumber: "",
      },
      AntiBPureRows: [],
      ColumnPreparationDate: "",
      ColumnPreparationDateRows: [],
      LabellingDate: "",
      LabellingDateRows: [],
      peptidePrepData: {
        BlockingPeptidePreparationDate: "",
        PeptideSupplier: "",
      },
      peptidePrepRows: [],

      IFCLotNumber: "",
      IFCLotNumberRows: [],
      IHCLotNumber: "",
      IHCLotNumberRows: [],

      isLoading: false,
      skeletonLoading: false,
    };
  }

  componentDidMount(): void {
    this.onInit();
  }

  async onInit() {
    this.setState({ isLoading: true })
    const products = await this.sp.web.lists.getById(this.props.ProductsListId).items();
    this.setState({ products: products, isLoading: false });
  }

  // Generic onChange
  onChange = (e: SelectChangeEvent) => {
    const { name, value } = e.target;

    this.setState({
      productId: value,
      AntiBPureRows: [],
      ColumnPreparationDateRows: [],
      LabellingDateRows: [],
      peptidePrepRows: []
    }, () => {
      this.setState({skeletonLoading: true})
      this.fetchProductDataFromLists()
    });
  };

  async fetchProductDataFromLists() {

    const lists = await this.sp.web.lists()

    const getItemsReq = lists.map(async (l: any) => {
      try {
        
        let items: any = await this.sp.web.lists.getById(l.Id).items
        .filter(`ProductId eq '${this.state.productId}'`).getPaged();
        

        // the results property will be an array of the items returned
        if (items.results?.length > 0) {
          console.log("getItemsReq - items:", items.results)
          items.results.forEach((r: any) => {
            if (r?.DevType === "AntibodyPurification") {
              const date = this.formatDate(r?.DateOfPurification)
              const AntibodyPurificationData = {
                DateOfPurification: date === "1/1/1970" ? "" : date,
                SerumNumber: r?.SerumNumber || "",
                ICANumber: r?.OData__x0023_ICA || "",
                ColumnNumber: r?.ColumnNumber || "",
                TotalQuantity: r?.Total_x0028_mg_x0029_CA || "",
                ExtraYieldCV: r?.ExtraYieldForStorageMG || "",
                LotNumber: r?.LotNumber || ""
              }
              this.setState({
                AntibodyPurificationData: AntibodyPurificationData,
                AntiBPureRows: [...this.state.AntiBPureRows, AntibodyPurificationData]
              })
            } else {
              const AntibodyPurificationData = {
                DateOfPurification:  "",
                SerumNumber: "",
                ICANumber:  "",
                ColumnNumber:  "",
                TotalQuantity:  "",
                ExtraYieldCV:  "",
                LotNumber:  ""
              }
              this.setState({
                AntibodyPurificationData: AntibodyPurificationData,
                AntiBPureRows: [...this.state.AntiBPureRows, AntibodyPurificationData]
              })
            }
            
            if (r?.DevType === "ColumnPreparation" || r?.DevType === "ColumnPreparationForFusionPeptide") {
              const ColPrepDate = this.formatDate(r?.DateOfColumnPreparation)
              this.setState({
                ColumnPreparationDate: ColPrepDate === "1/1/1970" ? "" : ColPrepDate,
                ColumnPreparationDateRows: [...this.state.ColumnPreparationDateRows, ColPrepDate === "1/1/1970" ? "" : ColPrepDate]
              })
            }
            if (r?.DevType === "AntibodyLabelling") {
              const LabellingDate = this.formatDate(r?.LabellingDate)
              this.setState({
                LabellingDate: LabellingDate === "1/1/1970" ? "" : LabellingDate,
                LabellingDateRows: [...this.state.LabellingDateRows, LabellingDate === "1/1/1970" ? "" : LabellingDate]
              })
            }
            if (r?.DevType === "BlockingPeptidePreparation" || r?.DevType === "FusionBlockingPeptidePreparation") { 
              const date = this.formatDate(r?.BlockingPeptidePreparationDate || r?.Date)
              const peptidePrepData = {
                BlockingPeptidePreparationDate: date === "1/1/1970" ? "" : date,
                PeptideSupplier: r?.Supplier || ""
              } 
              this.setState({
                peptidePrepData: peptidePrepData,
                peptidePrepRows: [...this.state.peptidePrepRows, peptidePrepData]
              })
            }

            if (r?.AppType === "IndirectFlowCytometry") {
              this.setState({
                IFCLotNumber: r?.LotNumber,
                IFCLotNumberRows: [...this.state.IFCLotNumberRows, r?.LotNumber]
              })
            }
            if (r?.AppType === "Immunohistochemistry") {
              
              this.setState({
                IHCLotNumber: r?.LotNumber,
                IHCLotNumberRows: [...this.state.IHCLotNumberRows, r?.LotNumber]
              })
            }

          })
        }
        if (items.hasNext) {
          // this will carry over the type specified in the original query for the results array
          items = await items.getNext();
        }
      } catch (err) {
        return null;
      }
    })

    Promise.all(getItemsReq).then(() => {
      this.setState({
        skeletonLoading: false
      })
    })
  }

  formatDate = (date: string) => {
    let realDate = new Date(date)
    let dd = String(realDate.getDate());
    let mm = String(realDate.getMonth() + 1); //January is 0!
    let yyyy = String(realDate.getFullYear());
    return dd + "/" + mm + "/" + yyyy;
  }

  renderTable() {
    const columns = [
      { field: 'DateOfPurification', headerName: 'Date Of Purification', width: 200 },
      { field: 'SerumNumber', headerName: 'Serum Number', width: 200 },
      { field: 'ICANumber', headerName: 'ICA Number', width: 200 },
      { field: 'ColumnNumber', headerName: 'Column Number', width: 200 },
      { field: 'ColumnPreparationDate', headerName: 'Column Preparation Date', width: 200 },
      { field: 'TotalQuantity', headerName: 'Total Quantity', width: 200 },
      { field: 'ExtraYieldCV', headerName: 'Extra Yield [C], (V)', width: 200 },
      { field: 'LotNumber', headerName: 'Lot Number', width: 200 },
      { field: 'LabellingDate', headerName: 'Labelling Date', width: 200 },
      { field: 'BlockingPeptidePreparationDate', headerName: 'Blocking Peptide Preparation Date', width: 250 },
      { field: 'PeptideSupplier', headerName: 'Peptide Supplier', width: 200 },
      { field: 'IFC', headerName: 'Indirect flow cytometry', width: 200 },
      { field: 'IHC', headerName: 'Immunohistochemistry', width: 200 },
    ];

    const rows = this.state.AntiBPureRows.map((row, index) => ({
      id: uuidv4(),
      DateOfPurification: row.DateOfPurification,
      SerumNumber: row.SerumNumber,
      ICANumber: row.ICANumber,
      ColumnNumber: row.ColumnNumber,
      TotalQuantity: row.TotalQuantity,
      ExtraYieldCV: row.ExtraYieldCV,
      LotNumber: row.LotNumber,
      LabellingDate: this.state.LabellingDateRows[index],
      ColumnPreparationDate: this.state.ColumnPreparationDateRows[index],
      BlockingPeptidePreparationDate: this.state.peptidePrepRows[index]?.BlockingPeptidePreparationDate,
      PeptideSupplier: this.state.peptidePrepRows[index]?.PeptideSupplier,
      IFC: this.state.IFCLotNumberRows[index],
      IHC: this.state.IHCLotNumberRows[index],
    }));

    return (
      <div style={{ width: '100%' }}>
        <DataGrid
          className={styles.dateGridToolBar}
          rows={rows}
          columns={columns}
          slots={{
            toolbar: GridToolbar,
            noRowsOverlay: this.state.skeletonLoading ? this.renderSkeletons : undefined
          }}
          sx={{
            height: rows.length === 0 ? '400px' : 'auto',
            '& .MuiDataGrid-virtualScroller::-webkit-scrollbar': {
              width: '0.4em',
            },
            '& .MuiDataGrid-virtualScroller::-webkit-scrollbar-track': {
              background: '#f1f1f1',
            },
            '& .MuiDataGrid-virtualScroller::-webkit-scrollbar-thumb': {
              backgroundColor: '#41a78e',
            },
            '& .MuiDataGrid-virtualScroller::-webkit-scrollbar-thumb:hover': {
              background: '#41a78e',
            },
            '& .MuiDataGrid-virtualScroller': {
              overflowY: 'hidden',
            }
          }}
        />
      </div>
    );
  }

  renderSkeletons() {
    const skeletonRows = Array.from(new Array(6)); // Number of skeletons
    return (
      <div style={{ height: 0, width: '100%' }}>
        {skeletonRows.map((_, index) => (
          <Skeleton key={index} variant="rectangular" height={30} style={{ marginBottom: 8 }} />
        ))}
      </div>
    );
  }
  

  public render(): React.ReactElement<IManagementControlProps> {
    const { Title } = this.props;

    return (
      <div className="EONewFormContainer">

        <div className="EOHeader">
          <div className="EOLogoContainer"></div>
          <div className="EOHeaderContainer">
            <span className="EOHeaderText">{Title}</span>
          </div>
        </div>
        {this.state.isLoading ? (
          <div className="SpinnerComp">
            <div className="loading-screen">
              <div className="loader-wrap">
                <span className="loader-animation"></span>
                <div className="loading-text">
                  <span className="letter">L</span>
                  <span className="letter">o</span>
                  <span className="letter">a</span>
                  <span className="letter">d</span>
                  <span className="letter">i</span>
                  <span className="letter">n</span>
                  <span className="letter">g</span>
                </div>
              </div>
            </div>
          </div>
        ) : (
          <section style={{ padding: "1em", textAlign: 'left' }}>
            
            <FormControl sx={{ m: 1, minWidth: 120, margin: 0, marginBottom: "1em" }} size="small">
              <InputLabel id="demo-select-small-label">Products</InputLabel>
              <Select
                labelId="demo-select-small-label"
                id="demo-select-small"
                value={this.state.productId}
                label="Product"
                name={"productId"}
                onChange={this.onChange}
              >
                <MenuItem value="">
                  <em>None</em>
                </MenuItem>
                {this.state.products.map((item) => (
                  <MenuItem key={uuidv4()} value={item.Id}>{item.ProductSerialNumber}</MenuItem>
                ))}
              </Select>
            </FormControl>
            <div className={styles.container}>
              {this.renderTable()}
            </div>
          </section>
        )}
      </div>

    );
  }
}
