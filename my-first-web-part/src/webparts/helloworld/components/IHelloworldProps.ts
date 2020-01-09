import { WebPartContext } from '@microsoft/sp-webpart-base';  

export interface IHelloworldProps {
  description: string;
  name:string;
  state:string;
  DropDownProp: string;  
  context: WebPartContext; 
}
