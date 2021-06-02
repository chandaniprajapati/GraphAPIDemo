import { MSGraphClient } from "@microsoft/sp-http";

export interface IGraphApiDemoProps {
  description: string;
  graphClient: MSGraphClient;
}
