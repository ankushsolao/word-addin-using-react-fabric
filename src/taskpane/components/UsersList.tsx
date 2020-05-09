import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn,
} from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { HttpService } from "../services/httpservice";
import {
    PrimaryButton, MessageBar, MessageBarType,
} from "office-ui-fabric-react";
//import { useBoolean } from '@uifabric/react-hooks';

const classNames = mergeStyleSets({
    fileIconHeaderIcon: {
        padding: 0,
        fontSize: '16px',
    },
    fileIconCell: {
        textAlign: 'center',
        selectors: {
            '&:before': {
                content: '.',
                display: 'inline-block',
                verticalAlign: 'middle',
                height: '100%',
                width: '0px',
                visibility: 'hidden',
            },
        },
    },
    fileIconImg: {
        verticalAlign: 'middle',
        maxHeight: '40px',
        maxWidth: '40px',
    },
    controlWrapper: {
        display: 'flex',
        flexWrap: 'wrap',
    },
    exampleToggle: {
        display: 'inline-block',
        marginBottom: '10px',
        marginRight: '30px',
    },
});
const controlStyles = {
    root: {
        margin: '0 30px 20px 0',
        maxWidth: '300px',
    },
};

export interface IDetailsListDocumentsExampleState {
    columns: IColumn[];
    items: IDocument[];
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
    selectedItem: IDocument,
    validationMessage: IValidation,
}

export interface IDocument {
    key: string;
    email: string;
    value: string;
    iconName: string;
    last_name: string;
    first_name: string;
}

export interface IValidation {
    isSuccess: boolean;
    message: string;
}

//const [teachingBubbleVisible, { toggle: toggleTeachingBubbleVisible }] = useBoolean(false);
export default class UsersList extends React.Component<{}, IDetailsListDocumentsExampleState> {

    private _selection: Selection;
    private _allItems: IDocument[];
    serv: HttpService;
    constructor(props: {}) {
        super(props);
        this.serv = new HttpService();
        const columns: IColumn[] = [
            {
                key: 'column1',
                name: 'Avatar',
                fieldName: 'avatar',
                minWidth: 50,
                maxWidth: 60,
                onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <img src={item.iconName} className={classNames.fileIconImg} />;
                },
            },
            {
                key: 'column2',
                name: 'Email',
                fieldName: 'email',
                minWidth: 160,
                maxWidth: 160,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'column3',
                name: 'First Name',
                fieldName: 'first_name',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                onRender: (item: IDocument) => {
                    return <span>{item.first_name}</span>;
                },
                isPadded: true,
            },
            {
                key: 'column4',
                name: 'Last Name',
                fieldName: 'last_name',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                isCollapsible: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.last_name}</span>;
                },
                isPadded: true,
            },
        ];


        this.state = {
            items: this._allItems,
            columns: columns,
            isModalSelection: false,
            isCompactMode: false,
            selectedItem: null,
            validationMessage: null,
        };
    }
    componentDidMount = () => {
        this.serv.getUsersList().then(resp => resp.data).then(data => {
            var userData = getUsersListData(data.data);
            this._allItems = userData;
            this.setState({
                items: userData,
            });
        }).catch(error => console.log("error :- ", error));
    };

    public render() {
        const { columns, isCompactMode, items } = this.state;
        return (
            <>
                {(items) ? (
                    <Fabric>
                        <div className={classNames.controlWrapper}>
                            <TextField label="Filter by Email:" onChange={this._onChangeText} styles={controlStyles} />
                        </div>
                        <div>
                            <PrimaryButton text="Validate" allowDisabledFocus onClick={this.handleValidate} />
                        </div>
                        <div>
                            {this.state.validationMessage ? (
                                this.state.validationMessage.isSuccess ? (
                                    <MessageBar
                                        messageBarType={MessageBarType.success}
                                    >
                                        {this.state.validationMessage.message}
                                    </MessageBar>
                                ) : (<MessageBar
                                    messageBarType={MessageBarType.error}
                                >
                                    {this.state.validationMessage.message}
                                </MessageBar>)
                            ) : null}
                        </div>
                        <DetailsList
                            items={items}
                            compact={isCompactMode}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            getKey={this._getKey}
                            setKey="none"
                            layoutMode={DetailsListLayoutMode.justified}
                            isHeaderVisible={true}
                            onItemInvoked={this.getUserData}
                        />
                    </Fabric>
                ) : null}
            </>
        );
    }
    // @ts-ignore
    public componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState) {
        if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
            this._selection.setAllSelected(false);
        }
    }
    // @ts-ignore
    private _getKey(item: any, index?: number): string {
        return item.key;
    }
    // @ts-ignore
    private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
        this.setState({
            items: text ? this._allItems.filter(i => i.email.toLowerCase().indexOf(text) > -1) : this._allItems,
        });
    };
    // @ts-ignore
    getUserData = async (item: any) => {
        this.setState({
            selectedItem: item,
        });
        console.log("getUserData ", item);

        return Word.run(async context => {
            var body = context.document.body;
            body.clear();
            var range = context.document.getSelection();
            var myArray = [["Email", "First Name", "Last Name"], [item.email, item.first_name, item.last_name]];
            var myTable = range.insertTable(2, 3, "Before", myArray);
            var myCC = myTable.insertContentControl();
            myCC.title = "myTableTitle";
            await context.sync();
        });
    }
    // @ts-ignore
    handleValidate = async () => {
        var selectedEmail = this.state.selectedItem.email;
        var documentEmail;
        return Word.run(async context => {
            var myTables = context.document.body.tables;
            context.load(myTables);
            await context.sync()
                .then(function () {
                    var myRows = myTables.items[0].rows;
                    context.load(myRows);
                    context.sync()
                        .then(function () {
                            documentEmail = myRows.items[1].values[0][0];
                            context.document.body.insertParagraph("Document Email:- " + documentEmail, Word.InsertLocation.end);
                            context.document.body.insertParagraph("Selected Email:- " + selectedEmail, Word.InsertLocation.end);
                            context.sync();                          
                        })
                });
            if (documentEmail === selectedEmail) {
                context.document.body.insertParagraph(selectedEmail, Word.InsertLocation.end);
            }        

            //  if(selectedEmail === documentEmail){
            //     this.setState({
            //         validationMessage: {
            //             isSuccess: true,
            //             message: "No Changes Found ",
            //         }
            //     });
            //  }
            //  else{
            //     this.setState({
            //         validationMessage: {
            //             isSuccess: false,
            //             message: "Data changed ",
            //         }
            //     });
            //  }
            await context.sync();
        });
    }

    // @ts-ignore
    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, items } = this.state;
        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
                this.setState({
                    announcedMessage: `${currColumn.name} is sorted ${
                        currColumn.isSortedDescending ? 'descending' : 'ascending'
                        }`,
                });
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });
        const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            items: newItems,
        });
    };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

function getUsersListData(data) {
    const items: IDocument[] = [];
    data.map((v, i) => (
        items.push({
            key: i.toString(),
            email: v.email,
            value: v.email,
            iconName: v.avatar,
            first_name: v.first_name,
            last_name: v.last_name,
        })
    ));
    return items;
}
