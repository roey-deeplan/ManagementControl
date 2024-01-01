export interface IManagementControlProps {
  Title: string;
  ProductsListId: string;
  context: any;

  //* Dev
  DevAntibodyPurificationId: string;
  DevBlockingPeptidepreparationId: string;
  DevFusionBlockingPeptidePreparationId: string;
  DevAntibodyLabellingId: string;
  DevColumnPreparationId: string;
  DevColumnPreparationForFusionPeptideId: string;
  //* QC
  QCAntibodyWesternBlotQcId: string;
  QcBlockingPeptideWesternBlotQcId: string;
  QcDirectFlowCytometryId: string;
  QcImmunohistochemistryId: string;
  //* Application
  AppIndirectFlowCytometryId: string;
  AppImmunohistochemistryId: string;
}
