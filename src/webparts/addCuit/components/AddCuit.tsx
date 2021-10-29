import * as React from 'react';
import styles from './AddCuit.module.scss';
import { IAddCuitProps } from './IAddCuitProps';
import {IAddCuitState} from './IAddCuitState';
import { IFolder } from "@pnp/sp/folders";
import { escape } from '@microsoft/sp-lodash-subset';
import { SearchBox, ISearchBoxStyles } from 'office-ui-fabric-react/lib/SearchBox';
import { sp } from "@pnp/sp/presets/all";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { PrimaryButton, DefaultButton, Stack, IStackTokens, TextField } from 'office-ui-fabric-react';
import { initializeIcons } from '@uifabric/icons';
import { Link, Text } from 'office-ui-fabric-react';

initializeIcons();

const dialogContentProps = {
  type: DialogType.normal,
  title: 'Cuit no encontrado',
  closeButtonAriaLabel: 'Close',
  subText: 'Desea crear una carpeta con su Cuit',
};
const existDialogContentProps ={
  type: DialogType.normal,
  title: 'Cuit existente',
  closeButtonAriaLabel: 'Close',
  subText: 'Su cuit ya existe',
}
const resultDialogContentProps ={
  type: DialogType.normal,
  title: 'Cuit creado exitosamente'
  
}

export default class AddCuit extends React.Component<IAddCuitProps, IAddCuitState> {

  constructor(props:IAddCuitProps){
    super(props);
    this.state = {
      spinner: false,
      cuitValue: "",
      dialog:false,
      existentCuit: false,
      resultDialog: false,
      result:"",
      folderURL:""
    }
  }

  protected async onInit(): Promise<void> {
  
    sp.setup(this.props.context);
  }
public componentDidMount() {
  console.log("contexto", window.location.origin )
/*  this.getFolder("124").then((carpetas)=>{
   console.log("Carpetas", carpetas)
 }).catch((err)=>{
   console.log("Error en getFolders", err)
 }) */
}

private  getFolder = (name:string): Promise<boolean> =>{
  this.setState({spinner:true})
 return sp.web.rootFolder.folders.getByName("Documentos compartidos").folders.getByName(name).get().then((carpeta)=>{
    console.log("carpeta encontrada", carpeta)
    this.setState({spinner:false, dialog: true,  existentCuit: true})
    return true
  }).catch((err)=>{
    console.log("Carpeta no encontrada", err)
    this.setState({spinner:false, dialog:true, existentCuit: false})
    return false
  })
}

private handleOnChangeCuit= (newValue:string):void =>{
  let onlyNums = newValue.replace(/[^0-9]/g, '');

    this.setState({cuitValue:onlyNums.length<12 ? onlyNums:this.state.cuitValue})


}

private createFolders = (name:string):Promise<any> =>{
  this.setState({spinner:true, dialog: false})
  return sp.web.rootFolder.folders.getByName("Documentos compartidos").folders.add(name)
  .then((res)=> {
    console.log("carpeta Cuit creada", res.data)
    let url=res.data.ServerRelativeUrl
      res.folder.addSubFolderUsingPath("Contratos Fija").then(async(contratosFolder)=>{
        console.log("carpeta Contratos Fija creada", contratosFolder)
        let numeroOsFolder = await contratosFolder.addSubFolderUsingPath("Numero Os")
        console.log("Carpeta Numero OS creada", numeroOsFolder)
        let contratoFirmadoFolder = await numeroOsFolder.addSubFolderUsingPath("Contrato Firmado")
        console.log("Carpeta Contrato Firmado creada", contratoFirmadoFolder)
        this.setState({spinner:false, resultDialog:true, folderURL:url})
     }).catch((err)=>{
       console.log("error en crear carpetea Contratos FIja", err)
       return err
     })
    res.folder.addSubFolderUsingPath("Documentación Facultativa "+name).then((resDocFacultativa)=>{
      console.log("Carpeta Documetacion Facultativa creada", resDocFacultativa)
    })
    res.folder.addSubFolderUsingPath("Documentación Móvil").then((resDocMovil)=>{
      console.log("Carpeta Documentación Móvil creada", resDocMovil);
      resDocMovil.addSubFolderUsingPath("Numero de SDS").then((resNumeroSDS)=>{
        console.log("Carpeta Numero SDS creada", resNumeroSDS);
      }).catch((err)=>{
        console.log("Error en crear carpeta Numero de SDS", err)
      })
    }).catch((err)=>{
      console.log("Error en crear carpeta Documentacion Movil", err)
    })
  }).catch((err)=>{
    console.log("error en crear carpeta Cuit: ", err)
    this.setState({spinner:false, resultDialog:true})
  })
}

private toggleHideDialog = ():void =>{
  this.setState({dialog:false, cuitValue:"", resultDialog:false, folderURL:""})
}

  public render(): React.ReactElement<IAddCuitProps> {
    return (
      <div className={ styles.addCuit }>
        {this.state.spinner?
        <Spinner size={SpinnerSize.large} />
        : 
        <div className={ styles.container }>
        <TextField
        value={this.state.cuitValue}
        onChange={(e,newValue) =>this.handleOnChangeCuit(newValue)}
        styles={{root:{width:200}}}
        placeholder="Crear Cuit"/>
         <PrimaryButton style={{ minWidth:20}} iconProps={{ iconName: 'Forward' }}  onClick={()=> this.getFolder(this.state.cuitValue)} allowDisabledFocus disabled={this.state.cuitValue==""} />
        
      <Dialog
        hidden={!this.state.dialog}
        onDismiss={this.toggleHideDialog}
        dialogContentProps={this.state.existentCuit?existDialogContentProps:dialogContentProps}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 350}  }
        }}
      >
        <DialogFooter>{this.state.existentCuit?
         <PrimaryButton onClick={this.toggleHideDialog} text="Ok" />
          :<>
          <PrimaryButton onClick={()=>this.createFolders(this.state.cuitValue)}  text="Si" />
          <DefaultButton onClick={this.toggleHideDialog} text="No" />
          </>}
        </DialogFooter>
      </Dialog>
      <Dialog
      hidden={!this.state.resultDialog}
      dialogContentProps={resultDialogContentProps}>
        <Text>
          Si desea ingresar a la carpeta creada haga clic {""}
        <Link href= {`${window.location.origin+this.state.folderURL}`} target="_blank " underline>
          aquí.
        </Link>
        </Text>
        <DialogFooter>
        <PrimaryButton onClick={this.toggleHideDialog} text="Ok" />
        </DialogFooter>
      </Dialog>
      </div>
        }
      </div>
    );
  }
}
