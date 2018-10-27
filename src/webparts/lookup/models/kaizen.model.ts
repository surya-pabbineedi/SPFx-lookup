import { Item } from 'sp-pnp-js/lib/sharepoint/items';

export class KaizenModel extends Item {
  public Id: number;
  public Title?: string;
  public Subject?: string;
  public Area_x0020__x002f__x0020_Locatio?: string;
  public Initial_x0020_Condition?: string;
  public Before?: string;
  public Benefits_x0020_Category?: string;
  public Process_x0020__x002f__x0020_Proj?: string;
  public Department_x0020__x002f__x0020_D?: string;
  public Solution_x0020_Description?: string;
  public After?: string;
  public Benefits_x0020_Description?: string;
  public Validated_x0020_By?: IPerson;
  public Approved_x0020_By?: IPerson;
  public Contact_x0020_Details?: string;
  public Team_x0020_Members?: string;
  public Implementation_x0020_Date?: Date;
  public Date_x0020_of_x0020_Completion?: Date;
  public Reference_x0023_?: string;
  public Standardization_x0020_Remarks?: string;
  public Author?: IPerson;
}

export interface IPerson {
  Id: number;
  Name: string;
  Title: string;
  EMail: string;
}