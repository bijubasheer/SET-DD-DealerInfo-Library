import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { AadHttpClient, HttpClient, IHttpClientOptions, HttpClientResponse, HttpClientConfiguration, MSGraphClientFactory, MSGraphClient, AadHttpClientFactory  } from '@microsoft/sp-http';
import { IUserItem } from "./IUserItem";
import { ISETDealerInfo } from "./ISETDealerInfo";

export interface IDealerInfoService {
  GetUserInfo():Promise<IUserItem>;
  GetDDInfo(nameId:string): Promise<ISETDealerInfo>;
}

export class DealerInfoService implements IDealerInfoService {
    
  //Create a ServiceKey which will be used to consume the service.
  public static readonly serviceKey: ServiceKey<IDealerInfoService> =
      ServiceKey.create<IDealerInfoService>('my-custom-app:IDealerInfoService', DealerInfoService);

  private _msGraphClientFactory: MSGraphClientFactory;
  private _aadHttpClientFactory: AadHttpClientFactory;
  private userItem: IUserItem = null;
  private dealerData: ISETDealerInfo = null;

  constructor(serviceScope: ServiceScope) {
      serviceScope.whenFinished(() => {
          this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
          this._aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);
      });
  }
  

  public async GetUserInfo():Promise<IUserItem> {
    if(!this.userItem)
    {
      console.log("userItem is not available. So need to fetch it");
      let client = await this._msGraphClientFactory.getClient();
      let response = await client.api('/me').get();
      this.userItem = response as  IUserItem;;
    }
    return this.userItem;
  }

  public async GetDDInfo(nameId:string):Promise<ISETDealerInfo> {
    if(!this.dealerData)
    {
      console.log("dealerData is not available. So need to fetch it");
      console.log("About to get Dealer Data for " + nameId);
      const getURL = "https://set.dev.api.jmfamily.com/dd-gwmenusvc-sys/v1/api/dealerContext?nameId=" + nameId;
      let setNum = "";
      const requestHeaders: Headers = new Headers();
      requestHeaders.append('Content-type', 'application/json');

      console.log("About to get SET Number");

      let client = await this._aadHttpClientFactory
      //.getClient('1148b2ca-eded-4ea7-9f1e-4cce4bd47109'); //CORPSTG1
      .getClient('b045f213-4f67-4081-a843-c49af374fab0');//NEW STG
      let response = await client.get(getURL, AadHttpClient.configurations.v1);
      this.dealerData = JSON.parse(await response.text());
      
    }    
    return Promise.resolve(this.dealerData);
  }

}

