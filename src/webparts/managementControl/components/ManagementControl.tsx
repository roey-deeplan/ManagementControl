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
  LotNumber: string;
  LotNumberRows: any[];
  AntibodyPurificationData: {
    DateOfPurification: Date | null,
    SerumNumber: string;
    ICANumber: string;
    ColumnNumber: string;
    TotalQuantity: string;
    ExtraYieldCV: string;
    LotNumber: string;
  };
  AntiBPureRows: any[];
  ColPrepData: {
    ColumnPreparationDate: Date | null,
    ColumnNumber: string,
  } | null;
  ColPrepDataRows: any[]
  Labelling: {
    LabellingDate: Date | null,
    LotNumber: string;
  };
  LabellingRows: any[];
  peptidePrepData: {
    BlockingPeptidePreparationDate: Date | null,
    PeptideSupplier: string,
    LotNumber: string;
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
      LotNumber: "",
      LotNumberRows: [],
      AntibodyPurificationData: {
        DateOfPurification: null,
        SerumNumber: "",
        ICANumber: "",
        ColumnNumber: "",
        TotalQuantity: "",
        ExtraYieldCV: "",
        LotNumber: "",
      },
      AntiBPureRows: [],
      ColPrepData: {
        ColumnPreparationDate: null,
        ColumnNumber: ""
      },
      ColPrepDataRows: [],
      Labelling: {
        LabellingDate: null,
        LotNumber: "",
      },
      LabellingRows: [],
      peptidePrepData: {
        BlockingPeptidePreparationDate: null,
        PeptideSupplier: "",
        LotNumber: "",
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

    // Reset the relevant state arrays and other related state variables before fetching new data.
    this.setState({
      productId: value,
      AntiBPureRows: [],
      ColPrepDataRows: [],
      LabellingRows: [],
      peptidePrepRows: [],
      IFCLotNumberRows: [],
      IHCLotNumberRows: [],
      LotNumberRows: [],
      // Reset any other relevant part of the state here.
    }, () => {
      this.setState({ skeletonLoading: true });
      this.fetchProductDataFromLists(); // Fetch new data based on the updated productId.
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
          items.results.forEach((r: any) => {
            // if (
            //   r?.DevType === "AntibodyLabelling" ||
            //   r?.DevType === "BlockingPeptidePreparation" ||
            //   r?.DevType === "FusionBlockingPeptidePreparation" ||
            //   r?.AppType === "Immunohistochemistry") {
            //   this.setState({
            //     LotNumber: r?.LotNumber,
            //     LotNumberRows: [...this.state.LotNumberRows, r?.LotNumber]
            //   })
            // }

            if (r?.DevType === "Antibody Purification") {
              const date = this.formatDate(r?.DateOfPurification)
              const AntibodyPurificationData = {
                DateOfPurification: date,
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
            }

            if (r?.DevType === "Column Preparation" || r?.DevType === "Column Preparation For Fusion Peptide") {
              const ColPrepDate = this.formatDate(r?.DateOfColumnPreparation);
              const ColPrepData = {
                ColumnPreparationDate: ColPrepDate,
                ColumnNumber: r?.ColumnNumber || "",
              };
              this.setState({
                ColPrepData: ColPrepData,
                ColPrepDataRows: [...this.state.ColPrepDataRows, ColPrepData]
              });
            }
            if (r?.DevType === "Antibody Labelling") {
              const Labelling = {
                LabellingDate: this.formatDate(r?.LabellingDate),
                LotNumber: r?.LotNumber || ""
              }
              this.setState({
                Labelling: Labelling,
                LabellingRows: [...this.state.LabellingRows, Labelling]
              })
            }
            if (r?.DevType === "Blocking Peptide Preparation" || r?.DevType === "Fusion Blocking Peptide Preparation") {
              const date = this.formatDate(r?.BlockingPeptidePreparationDate || r?.Date)
              const peptidePrepData = {
                BlockingPeptidePreparationDate: date,
                PeptideSupplier: r?.Supplier || "",
                LotNumber: r?.LotNumber || ""
              }
              this.setState({
                peptidePrepData: peptidePrepData,
                peptidePrepRows: [...this.state.peptidePrepRows, peptidePrepData]
              })
            }

            if (r?.AppType === "Indirect Flow Cytometry") {
              this.setState({
                IFCLotNumber: r?.LotNumber,
                IFCLotNumberRows: [...this.state.IFCLotNumberRows, r?.LotNumber]
              })
            }
            if (r?.AppType === "Immunohistochemistry") {
              this.setState({
                IHCLotNumber: r?.LotNumberOrProductionDate,
                IHCLotNumberRows: [...this.state.IHCLotNumberRows, r?.LotNumberOrProductionDate]
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
    const timestamp = Date.parse(date);
    if (!isNaN(timestamp)) {
      return new Date(timestamp);
    }
    return null; // Or return a default date, depending on your needs
  }

  renderTable() {
    const columns = [
      { field: 'DateOfPurification', headerName: 'Date Of Purification', width: 200, type: 'date' },
      { field: 'SerumNumber', headerName: 'Serum Number', width: 200 },
      { field: 'ICANumber', headerName: 'ICA Number', width: 200 },
      { field: 'ColumnNumber', headerName: 'Column Number', width: 200 },
      { field: 'TotalQuantity', headerName: 'Total Quantity', width: 200 },
      { field: 'ExtraYieldCV', headerName: 'Extra Yield (mg)', width: 200 },
      { field: 'LotNumber', headerName: 'Lot Number', width: 200 },
      { field: 'ColumnPreparationDate', headerName: 'Column Preparation Date', width: 200, type: 'date' },
      { field: 'LabellingDate', headerName: 'Labelling Date', width: 200, type: 'date' },
      { field: 'BlockingPeptidePreparationDate', headerName: 'Blocking Peptide Preparation Date', width: 250, type: 'date' },
      { field: 'PeptideSupplier', headerName: 'Peptide Supplier', width: 200 },
      { field: 'IFC', headerName: 'Indirect flow cytometry', width: 200 },
      { field: 'IHC', headerName: 'Immunohistochemistry', width: 200 },
    ];

    // Combine all rows into a single array
    const combinedRows = [
      // ...this.state.LotNumberRows.map(lotNumber => ({
      //   LotNumber: lotNumber, id: uuidv4()
      // })),
      ...this.state.AntiBPureRows.map(row => ({ ...row, id: uuidv4() })),
      ...this.state.ColPrepDataRows.map(row => ({ ...row, id: uuidv4() })),
      ...this.state.LabellingRows.map(row => ({ ...row, id: uuidv4() })),
      ...this.state.peptidePrepRows.map(row => ({ ...row, id: uuidv4() })),
      ...this.state.IFCLotNumberRows.map(lotNumber => ({
        IFC: lotNumber, id: uuidv4()
      })),
      ...this.state.IHCLotNumberRows.map(lotNumber => ({
        IHC: lotNumber, id: uuidv4(),
        LotNumber: lotNumber
      })),
      // Add other arrays similarly
    ];

    return (
      <div style={{ width: '100%' }}>
        <DataGrid
          className={styles.dateGridToolBar}
          rows={combinedRows}
          columns={columns}
          slots={{
            toolbar: GridToolbar,
            noRowsOverlay: this.state.skeletonLoading ? this.renderSkeletons : undefined
          }}
          sx={{
            height: combinedRows.length === 0 ? '400px' : 'auto',
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
