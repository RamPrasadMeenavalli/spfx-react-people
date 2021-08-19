import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPeopleExplorerProps {
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateTitle: (value:string) => void;

  people: any[];
  updatePeople: (values:any[]) => void;

  template: string;
}
