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

  isLoading: boolean;
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

      isLoading: false,
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
      this.fetchProductDataFromLists()
    });
  };

  async fetchProductDataFromLists() {

    const lists = await this.sp.web.lists()

    const getItemsReq = lists.map(async (l: any) => {
      try {
        const items: any = await this.sp.web.lists.getById(l.Id).items.filter(`ProductId eq '${this.state.productId}'`).getPaged();
        if (items.results?.length) {
        }
        console.log("getItemsReq - items:", items)
        items.results.forEach((r: any) => {
          if (r?.DevType === "AntibodyPurification") {

            const AntibodyPurificationData = {
              DateOfPurification: this.formatDate(r?.DateOfPurification) || "-",
              SerumNumber: r?.SerumNumber || "-",
              ICANumber: r?.OData__x0023_ICA || "-",
              ColumnNumber: r?.ColumnNumber || "-",
              TotalQuantity: r?.Total_x0028_mg_x0029_CA || "-",
              ExtraYieldCV: r?.ExtraYieldForStorageMG || "-",
              LotNumber: r?.LotNumber || "-"
            }
            this.setState({
              AntibodyPurificationData: AntibodyPurificationData,
              AntiBPureRows: [...this.state.AntiBPureRows, AntibodyPurificationData]
            })
          }
          if (r?.DevType === "ColumnPreparation" || r?.DevType === "ColumnPreparationForFusionPeptide") {
            const ColPrepDate = this.formatDate(r?.DateOfColumnPreparation) || "-"
            this.setState({
              ColumnPreparationDate: ColPrepDate,
              ColumnPreparationDateRows: [...this.state.ColumnPreparationDateRows, ColPrepDate]
            })
          }
          if (r?.DevType === "AntibodyLabelling") {
            const LabellingDate = this.formatDate(r?.LabellingDate) || "-"
            this.setState({
              LabellingDate: LabellingDate,
              LabellingDateRows: [...this.state.LabellingDateRows, LabellingDate]
            })
          }
          if (r?.DevType === "BlockingPeptidePreparation") {
            const peptidePrepData = {
              BlockingPeptidePreparationDate: this.formatDate(r?.BlockingPeptidePreparationDate) || "-",
              PeptideSupplier: r?.Supplier || "-"
            }

            this.setState({
              peptidePrepData: peptidePrepData,
              peptidePrepRows: [...this.state.peptidePrepRows, peptidePrepData]
            })
          }
        })
      } catch (err) {
        return null;
      }
    })

    Promise.all(getItemsReq).then(() => {
      //console.log("fetchProductDataFromLists - newItems:", newItems)
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
      { field: 'TotalQuantity', headerName: 'Total Quantity', width: 200 },
      { field: 'ExtraYieldCV', headerName: 'Extra Yield [C], (V)', width: 200 },
      { field: 'LotNumber', headerName: 'Lot Number', width: 200 },
      { field: 'LabellingDate', headerName: 'Labelling Date', width: 200 },
      { field: 'ColumnPreparationDate', headerName: 'Column Preparation Date', width: 200 },
      { field: 'BlockingPeptidePreparationDate', headerName: 'Blocking Peptide Preparation Date', width: 250 },
      { field: 'PeptideSupplier', headerName: 'Peptide Supplier', width: 200 },
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
    }));

    return (
      <div style={{ height: 400, width: '100%' }}>
        <DataGrid
          className={styles.dateGridToolBar}
          rows={rows}
          columns={columns}
          slots={{
            toolbar: GridToolbar,
          }}
          sx={{
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
          }}
        />
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
                  <MenuItem key={uuidv4()} value={item.Id}>{item.ProductName.Description}</MenuItem>
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
