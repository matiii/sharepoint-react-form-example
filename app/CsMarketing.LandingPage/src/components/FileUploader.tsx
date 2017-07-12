import * as React from 'react';
import * as Dropzone from 'react-dropzone'
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { List } from 'office-ui-fabric-react/lib/List';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import * as update from 'immutability-helper';

export interface FileUploaderProps {
    onUploadedFilesChange: (files: { file: File, base64: string }[]) => void;
}

export interface FileUploaderState {
    selectedFiles: { file: File, percentProgress: number, description: string }[];
}

export interface SelectedFilesProps {
    files: { file: File, percentProgress: number, description: string }[];

    onRemoveFile(index: number);
}

const style: React.CSSProperties = {
    width: '600px',
    height: '350px',
    border: '1px solid #c8c8c8',
    cursor: 'pointer'
}

const childrenStyle: React.CSSProperties = {
    fontFamily: '"Segoe UI WestEuropean","Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,"Helvetica Neue",sans-serif;',
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    flexDirection: 'column',
    width: '100%',
    height: '100%',
    fontSize: '16px'
}

const spanStyle: React.CSSProperties = {
    marginLeft: '7px'
}

const errorMsg = 'Please remove file and upload again.';

const SelectedFiles: React.StatelessComponent<SelectedFilesProps> = (props: SelectedFilesProps) => {

    const excelExtensions = ['.xls', '.xlt', '.xml', '.xlsx', '.xlsm', '.xltx', '.xltm'];
    const wordExtensions = ['.doc', '.dot', '.wbk', '.docx', '.docm', '.dotx', '.dotm', '.docb'];
    const powerPointExtensions = ['.ppt', '.pot', '.pps', '.pptx', '.pptm', '.potx', '.potm', '.ppam', '.ppsx', '.ppsm', '.sldx', '.sldm'];
    const photoExtensions = ['.png', '.jpg', '.jpeg'];

    const logoStyle: React.CSSProperties = {
        marginRight: '25px',
        fontSize: '25px',
        position: 'relative',
        top: '-15px',
        left: '9px'
    };

    const progressStyle: React.CSSProperties = {
        width: '500px',
        display: 'inline-block'
    }

    const containerStyle: React.CSSProperties = {
        marginBottom: '15px'
    };

    const itemStyle: React.CSSProperties = {
        width: '600px'
    }

    return (
        <div style={containerStyle}>
            <Label>Selected files: </Label>

            <List
                items={props.files}
                onRenderCell={(item, index) => {

                    const resolveIcon = (name: string) => {
                        if (excelExtensions.filter(x => name.toLowerCase().indexOf(x) > -1).length > 0) {
                            return 'ExcelLogo';
                        }

                        if (wordExtensions.filter(x => name.toLowerCase().indexOf(x) > -1).length > 0) {
                            return 'WordLogo';
                        }

                        if (powerPointExtensions.filter(x => name.toLowerCase().indexOf(x) > -1).length > 0) {
                            return 'PowerPointLogo';
                        }

                        if (photoExtensions.filter(x => name.toLowerCase().indexOf(x) > -1).length > 0) {
                            return 'Photo2';
                        }

                        if (name.toLowerCase().indexOf('pdf') > -1) {
                            return 'PDF';
                        }

                        return 'OpenFile';
                    };

                    let logo = resolveIcon(item.file.name);

                    return (
                        <div style={itemStyle}>
                            <i style={logoStyle} className={`ms-Icon ms-Icon--${logo}`} aria-hidden="true"></i>
                            <div style={progressStyle}>
                                <ProgressIndicator
                                    label={item.file.name}
                                    description={item.description}
                                    percentComplete={item.percentProgress} />
                            </div>
                            {(item.percentProgress === 1 || item.description.indexOf(errorMsg) > -1) && <IconButton onClick={() => props.onRemoveFile(index)} style={{ top: '16px' }} iconProps={{ iconName: 'Cancel' }} title='Remove' />}
                        </div>
                    );
                }} />


        </div>
    );
}


export class FileUploader extends React.Component<FileUploaderProps, FileUploaderState>{

    private readonly uploadedFiles: { file: File, base64: string }[] = [];

    constructor(props: FileUploaderProps, context: any) {
        super(props, context);
        this.state = { selectedFiles: [] };
    }

    @autobind
    private onDropHandler(acceptedFiles: File[], rejectedFiles: FileList) {
        let lengthBefore = this.state.selectedFiles.length;

        this.setState(update(this.state, { selectedFiles: { $push: acceptedFiles.map(x => ({ file: x, percentProgress: 0, description: 'Start uploading.' })) } }));

        let uploadFile = (file: File, index: number) => {
            var reader = new FileReader();
            
            reader.onprogress = (ev: ProgressEvent) => {
                let percent = ev.loaded / ev.total;
                this.setState(update(this.state, { selectedFiles: { [index]: { percentProgress: { $set: percent }, descritpion: { $set: `${percent * 100} / 100` } } } }));
            };

            reader.onloadend = (ev: ProgressEvent) => {
                if (reader.error) {
                    this.setState(update(this.state, { selectedFiles: { [index]: { percentProgress: { $set: 0 }, description: { $set: `${reader.error.name} ${errorMsg}` } } } }));
                } else {
                    this.setState(update(this.state, { selectedFiles: { [index]: { percentProgress: { $set: 1 }, description: { $set: 'Complete' } } } }));

                    let base64 = (reader.result as string).replace(/^.*;base64,/, '');

                    this.uploadedFiles.push({ file, base64 });

                    this.notifyUploadedFiles();
                }

                index++;

                if (index - lengthBefore < acceptedFiles.length) {
                    uploadFile(acceptedFiles[index - lengthBefore], index);
                }
            };

            reader.readAsDataURL(file);
        };

        uploadFile(acceptedFiles[0], lengthBefore);
    }

    @autobind
    private removeFileHandler(index: number) {
        this.setState(update(this.state, { selectedFiles: { $splice: [[index, 1]] } }));

        this.uploadedFiles.splice(index, 1);

        this.notifyUploadedFiles();
    }

    @autobind
    private notifyUploadedFiles() {
        this.props.onUploadedFilesChange([...this.uploadedFiles]);
    }

    render() {
        return (
            <div>
                {this.state.selectedFiles.length > 0 && <SelectedFiles onRemoveFile={this.removeFileHandler} files={this.state.selectedFiles} />}

                <Dropzone style={style} className='cs-file-uploader' onDrop={this.onDropHandler}>
                    <div style={childrenStyle}>
                        <div>
                            <i className="ms-Icon ms-Icon--Attach" aria-hidden="true"></i>
                            <span style={spanStyle}>Attach files</span>
                        </div>
                    </div>
                </Dropzone>
            </div>
        );
    }
} 