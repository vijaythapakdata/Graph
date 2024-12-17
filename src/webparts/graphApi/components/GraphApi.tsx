import * as React from 'react';
// import styles from './GraphApi.module.scss';
import type { IGraphApiProps } from './IGraphApiProps';
import { IGraphApiState } from './IGraphApiState';
import {GraphError, ResponseType} from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { Link, Persona, PersonaSize } from '@fluentui/react';
export default class GraphApi extends React.Component<IGraphApiProps, IGraphApiState> {
  constructor(props:any){
    super(props);
    this.state={
      name:"",
      email:"",
      phone:"",
      image:""
    }
  }

  //Get Emal

  private _renderEmail=():JSX.Element=>{
    if(this.state.email){
      return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>
    }
    else{
      return<div/>
    }
  }
  //Get Phone 
  private _renderPhome=():JSX.Element=>{
    if(this.state.email){
      return <Link href={`tel:${this.state.phone}`}>{this.state.phone}</Link>
    }
    else{
      return<div/>
    }
  }

  public componentDidMount(): void {
    this.props.graphClient.api('me')
    .get((error:GraphError,user:MicrosoftGraph.User)=>{
      this.setState({
        name:user.displayName,
        email:user.mail,
        phone:user.businessPhones?.[0]
      });
    });
    this.props.graphClient.api("me/photo/$value")
    .responseType(ResponseType.BLOB).get((error:GraphError,photoresponse:Blob)=>{
      const bloburl=window.URL.createObjectURL(photoresponse);
      this.setState({image:bloburl});
    })
    
  }
  public render(): React.ReactElement<IGraphApiProps> {
   

    return (
   <>
   <Persona primaryText={this.state.name}
   
   secondaryText={this.state.email}
   onRenderSecondaryText={this._renderEmail}
   tertiaryText={this.state.phone}
   onRenderTertiaryText={this._renderPhome}
   imageUrl={this.state.image}
   size={PersonaSize.size100}/>
   </>
    );
  }
}
