import * as React from 'react';
import styles from './ManagementControl.module.scss';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select, { SelectChangeEvent } from '@mui/material/Select';
import type { IManagementControlProps } from './IManagementControlProps';
import getSP from "../PnPjsConfig";
import { v4 as uuidv4 } from 'uuid';
export interface IManagementControlState {
  products: any[];
  product: string;
}

export default class ManagementControl extends React.Component<IManagementControlProps, IManagementControlState> {
  sp = getSP(this.props.context);

  constructor(props: IManagementControlProps) {
    super(props);

    this.state = {
      products: [],
      product: ""
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
    this.setState({ product: e.target.value as string });
  };
  
  async fetchProductDataFromLists() {

  }

  public render(): React.ReactElement<IManagementControlProps> {
    const { Title } = this.props;

    return (
      <section>
        <h1>{Title}</h1>
        <FormControl sx={{ m: 1, minWidth: 120 }} size="small">
          <InputLabel id="demo-select-small-label">Products</InputLabel>
          <Select
            labelId="demo-select-small-label"
            id="demo-select-small"
            value={this.state.product}
            label="Product"
            onChange={this.onChange}
          >
            <MenuItem value="">
              <em>None</em>
            </MenuItem>
            {this.state.products.map((item, index) => (
              <MenuItem key={uuidv4()} value={item.ProductName.Description}>{item.ProductName.Description}</MenuItem>
            ))}
          </Select>
        </FormControl>
      </section>
    );
  }
}
