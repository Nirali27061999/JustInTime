import * as React from 'react';
import { Dropdown, DatePicker, DefaultButton, PrimaryButton, TextField, Dialog, DialogType, DialogFooter, Label } from 'office-ui-fabric-react';
import { IJustInTimeWebpartProps } from './IJustInTimeWebpartProps';
import 'office-ui-fabric-react/dist/css/fabric.css';
import styles from './JustInTimeWebpart.module.scss';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
// import { sp } from "@pnp/sp/presets/all";
import { SPFx, spfi } from "@pnp/sp";
import { NumericFormat } from 'react-number-format';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/items/get-all";
import * as moment from 'moment';
//  import './JustinTime.scss';
  import '../assets/custom.css';
 
// import 'office-ui-fabric-react/dist/css/fabric.css';
// import { Field } from "@pnp/sp/fields";
//   import * as moment from 'moment';
export default class SiteCollectionForList extends React.Component<IJustInTimeWebpartProps, any> {
    constructor(props: IJustInTimeWebpartProps) {
        super(props);
        const currentDate = new Date();
        this.state = {
            siteUrlOptions: [],
            siteUrl: '',
            groupNameOptions: [],
            groupName: '',
            userName: '',
            // addDate: null,
            addDate:currentDate,
            removeDate: null,
            removeDate1:null,
            ApprovalUser: '',
            // for group ID 
            groupIdoptions: [],
            groupID: '',
            approvalOptions: [],
            selectedApprovalKey: '',
            reason:'',
            RevokeDays: 0,
            dialogSubText : '',
            // isadddate:false,
            // isremovedate:false,

            // For Validation
            siteUrlErrorMessage: '',
            groupNameErrorMessage: '',
            userNameErrorMessage: '',
            addDateErrorMessage: '',
            removeDateErrorMessage: '',
            ApprovalUserErrorMessage: '',
            reasonErrorMessage: '',
            isDialogVisible: false, // New state variable for the dialog
            isSpinnerLoader: false,
            

        };
    }

    async componentDidMount() {
        SPComponentLoader.loadCss(`${this.props.context.pageContext.web.absoluteUrl}/SiteAssets/CustomCssInject.css`);
        await this.loadSiteUrls();
        await this.loadApprovalOptions();
    }
    loadSiteUrls = async () => {
        // const siteCollectionsUrl = this.props.context.pageContext.web.absoluteUrl + "/_api/search/query?querytext='contentclass:STS_Site'";

        // SubSite Approch
        const wholeSite = this.props.context.pageContext.web.absoluteUrl;
  const urlObject = new URL(wholeSite);
const rootSiteUrl = `${urlObject.protocol}//${urlObject.hostname}`;
  const siteCollectionsUrl = rootSiteUrl + "/_api/search/query?querytext='contentclass:STS_Site%20contentclass:STS_Web'&selectproperties='Title,Path'&rowlimit=500";

        const response = await fetch(siteCollectionsUrl, {
            headers: {
                Accept: 'application/json;odata=nometadata',
            },
        });

        if (response.ok) {
            const data = await response.json();
            const relevantResults = data.PrimaryQueryResult.RelevantResults;
            const siteCollections = [];

            if (relevantResults.Table && relevantResults.Table.Rows && relevantResults.Table.Rows.length > 0) {
                const rows = relevantResults.Table.Rows;
                const pathIndex = rows[0].Cells.findIndex((cell: any) => cell.Key === 'Path');
                // const titleIndex = rows[0].Cells.findIndex((cell: any) => cell.Key === 'Title');
                const titleIndex = rows[0].Cells.findIndex((cell: any) => cell.Key === 'Path');

                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    const path = row.Cells[pathIndex].Value;
                    const title = row.Cells[titleIndex].Value;

                    siteCollections.push({
                        key: path,
                        text: title,
                    });
                }
            }

            this.setState({
                siteUrlOptions: siteCollections,
              //  siteUrl: siteCollections.length > 0 ? siteCollections[0].key : '',
            });
        } else {
            console.error("Failed to fetch site URLs");
        }
    }

    loadGroupNames = async (siteUrl: string) => {
        const requestUrl = siteUrl + "/_api/web/sitegroups";
        const requestOptions = {
            headers: {
                Accept: "application/json;odata=verbose"
            }
        };

        try {
            const response = await fetch(requestUrl, requestOptions);
            const data = await response.json();

            if (response.ok) {
                const groups = data.d.results;
                const groupNameOptions = groups.map((group: any) => ({ key: group.Title, text: group.Title }));
                this.setState({ groupNameOptions });

                const selectedGroup = this.state.groupName;
                const selectedGroupData = groups.find((group: any) => group.Title === selectedGroup);
                const groupIdoptions = selectedGroupData ? [{ key: selectedGroupData.Id, text: selectedGroupData.Id }] : [];
                this.setState({ groupIdoptions });
                console.log(data,"loadedgroupdata");
            } else {
                console.error("Failed to fetch site groups:", data.error);
            }
        } catch (error) {
            console.error("An error occurred while fetching site groups:", error);
        }
    };
    loadApprovalOptions = async () => {
        try {
            const listName = "Jittest";
            const choiceFieldInternalName = "Approval";
            const siteUrl = 'https://imrchusky.sharepoint.com/sites/test'; // Replace with your site URL
            const endpoint = `${siteUrl}/_api/web/lists/GetByTitle('${listName}')/fields?$filter=EntityPropertyName eq '${choiceFieldInternalName}'`;
            const response = await fetch(endpoint, {
                headers: {
                    accept: "application/json;odata=verbose",
                },
            });

            if (response.ok) {
                const data = await response.json();
                console.log(data); // Check the entire structure of the data returned
                const choices = data.d.results[0].Choices;
                console.log(choices); // Check the type and structure of the choices
                if (Array.isArray(choices.results)) {
                    const approvalOptions = choices.results.map((choice: string, index: number) => ({ key: index.toString(), text: choice }));
                    this.setState({ approvalOptions });
                } else {
                    console.error("Choices is not an array");
                }
            } else {
                console.error("Failed to fetch Approval options");
            }
        } catch (error) {
            console.error("Error fetching Approval options: ", error);
        }
    };


    public numberFormatTextChangeEvent = (controlName: string, e: any) => {
        const { floatValue, formattedValue } = e;
        if (floatValue !== undefined) {
            this.setState({ RevokeDays : floatValue });
          }
      }


    saveData = async () => {
        // const { siteUrl, groupName, userName, addDate, removeDate, selectedApprovalKey, ApprovalUser } = this.state;
        const { siteUrl, groupName, userName, addDate, removeDate, ApprovalUser,reason,removeDate1,RevokeDays,dialogSubText } = this.state;
        console.log(this.state.ApprovalUser,"AppUser");
        this.setState({ isSpinnerLoader: true });
        // For Validation

        let isFormValid = true;

        if (!siteUrl) {
            this.setState({ siteUrlErrorMessage: 'Site URL is Required' });
            isFormValid = false;
        } else {
            this.setState({ siteUrlErrorMessage: '' });
        }

        if (!groupName) {
            this.setState({ groupNameErrorMessage: 'Group Name is Required' });
            isFormValid = false;
        } else {
            this.setState({ groupNameErrorMessage: '' });
        }

        if (!userName) {
            this.setState({ userNameErrorMessage: 'User Name is Required' });
            isFormValid = false;
        } else {
            this.setState({ userNameErrorMessage: '' });
        }

        if (ApprovalUser.length === 0) {
            this.setState({ ApprovalUserErrorMessage: 'Approver Name is Required' });
            isFormValid = false;
        } else {
            this.setState({ ApprovalUserErrorMessage: '' });
            if(ApprovalUser[0].secondaryText !== this.props.context.pageContext.user.email)
        {
            this.setState({ dialogSubText: 'Permission will be assigned after approval' });
            if (!reason) {
                this.setState({ reasonErrorMessage: 'Remarks is Required' });
                isFormValid = false;
            } else {
                this.setState({ reasonErrorMessage: '' });
            }
        }
        else{
            this.setState({ dialogSubText: 'Permission is assigned to this user sucessfully' });
        }
        }
        
        if (!RevokeDays) {
            this.setState({ removeDateErrorMessage: 'Revoke Days is Required' });
            isFormValid = false;
        } else {
            if(RevokeDays > 30)
            {
                this.setState({ removeDateErrorMessage: 'Revoke Days should not be greater then 30' });
                isFormValid = false;
            }
            else{
                this.setState({ removeDateErrorMessage: '' });
            }
           
        }
        
        
 
        if (!isFormValid) {
            this.setState({ isSpinnerLoader: false });
            return;
        }
       
        const sp = spfi().using(SPFx(this.props.context));
        // const ad=  moment(addDate).utc().format('DD-MM-YYYY');
        // const rd=  moment(removeDate).utc().format('DD-MM-YYYY');
        const a = this.state.groupIdoptions[0].key;
        const daysDifference = this.calculateDateDifference();
         const today = moment();
         const expirationDate = moment(this.state.removeDate);

        // // Calculate the difference in days from today to the expiration date
         const daysDifferenceFromToday = expirationDate.diff(today, 'days');

        if (daysDifferenceFromToday > 20) {
            alert("Expiration date should not be more than 20 days from the current date.");
            return;
        }

        const list = await sp.web.lists.getByTitle("Jittest").items.add({
            'Title': 'Abhi',
            'SiteUrl': siteUrl,
            'GroupName': groupName,
            'UserNameId': userName[0].id,
            'AddDate': addDate.toISOString(),
         
            'GroupId': a.toString(),
            // 'Approval': this.state.approvalOptions[parseInt(selectedApprovalKey)].text, // Assuming selectedApprovalKey is the index of the selected option,
            'ApprovalUserId': ApprovalUser[0].id,
            'Expires': RevokeDays.toString(),
            'Reason':reason
        });
        // alert("Data Insert Successfull");
        console.log(list);
        this.setState({ isDialogVisible: true }); // Set the state to show the dialog
        this.resetForm();
    }
    dismissDialog = () => {
      this.setState({ isDialogVisible: false });
        // Redirect to the homepage
     window.location.href = this.props.context.pageContext.web.absoluteUrl; // Replace with your actual homepage URL
  }
  
    resetForm = () => {
        this.setState({
            siteUrl: '',
            groupName: '',
            addDate: new Date(),
            removeDate: null,
            groupID: '',
            selectedApprovalKey: '', // Reset selectedApprovalKey
            userName: [], // Reset userName,
            ApprovalUser: [],
            reason:''

        });
        this.setState({ isSpinnerLoader: false });
    }

    calculateDateDifference = () => {
        const { addDate, removeDate } = this.state;
        if (addDate && removeDate) {
            const startDate = moment(addDate, 'DD-MM-YYYY');
            const endDate = moment(removeDate, 'DD-MM-YYYY');
            const duration = moment.duration(endDate.diff(startDate));
            const daysDifference = Math.floor(duration.asDays());
            return daysDifference;
        }
        return 0;
    };
    handleAddDateChange = (date: Date | null | undefined): void => {
         this.setState({ addDate: date });
    };

    handleApproverNameChange = (items: any) => {
        // Update state for the first control
        this.setState({ ApprovalUser: items });
    
        // Reset the value in the second control
        this.setState({ reason: null });
      }

    handleRemoveDateChange = (date: Date | null | undefined): void => {
        const selectedDate = moment(date);
        const startDate = moment(this.state.addDate);

        if (selectedDate.isSame(startDate, 'day')) {
          alert("Expiry date should not be the same as the start date.");
          this.setState({ removeDate: null });
          return;
      }
        const nextTwentyDays = moment().add(20, 'days');
        const beforeTwentyDays = moment().subtract(20, 'days');

        // Check if the selected date is more than 20 days before today's date
        if (selectedDate.isBefore(beforeTwentyDays, 'day')) {
            alert("Please select a date not more than 20 days before today's date.");
            this.setState({ removeDate: null });
            return;
        }

        // Check if the selected date is more than 20 days from today's date
        if (selectedDate.isAfter(nextTwentyDays, 'day')) {
            alert("Please select a date within the next 20 days from today's date.");
            this.setState({ removeDate: null });
            return;
        }
   // Add 1 day to the selected date
    const updatedRemoveDate = selectedDate.add(1, 'day').toDate();
        // Set the state only if the selected date passes the validation
        //  this.setState({ removeDate: date });
         this.setState({ removeDate1:date });
         this.setState({ removeDate: updatedRemoveDate });
       
    };

    render() {
      // const { siteUrl, groupName, userName, addDate, removeDate, ApprovalUser, reason } = this.state;
      const {   siteUrlErrorMessage, groupNameErrorMessage, userNameErrorMessage,ApprovalUserErrorMessage,reasonErrorMessage,addDateErrorMessage,removeDateErrorMessage,dialogSubText} = this.state;
      const minDate = new Date();
   const nextdays:any= minDate.setDate(minDate.getDate() + 1);
        return (

<>
{this.state.isSpinnerLoader ?
    <div className={styles.overlay}>
     <div className={styles.loader}>
       <img src={require('../assets/images/gear.gif')} alt="Loading..." />
     </div>
   </div>
   : undefined}
<div dir="ltr" className={`ms-Grid ${styles.tabWrapper}`}>
<div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
    <h1 style={{ textAlign: 'center' }}>Just In Time Access Control</h1>
</div>
<div className="ms-Grid-row">
    {/* Site URL */}
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
        <div className={`customControl ${styles.formControl}`}>
            <Label className='customLabel'>Site URL<span style={{ color: 'red' }}> * </span></Label>
            <Dropdown
                //label="Site URL"
                options={this.state.siteUrlOptions}
                selectedKey={this.state.siteUrl}
                onChange={async (e, option) => {
                    if (option) {
                        const selectedKey = option.key.toString(); // Convert the key to a string
                        await this.setState({ siteUrl: selectedKey });
                        await this.loadGroupNames(selectedKey);
                    }
                }}
                // errorMessage={siteUrl ? '' : 'This field is required.'}
            />
            {siteUrlErrorMessage && <span style={{ color: 'red' }}>{siteUrlErrorMessage}</span>}
        </div>
    </div>

    {/* Group Name */}
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
        <div className={`customControl ${styles.formControl}`}>
            <Label className='customLabel'>Group Name<span style={{ color: 'red' }}> * </span></Label>
            <Dropdown
                //label="Group Name"
                options={this.state.groupNameOptions}
                selectedKey={this.state.groupName}
                onChange={(e, option) => {
                    if (option) {
                        this.setState({ groupName: option.key as string }, () => {
                            const selectedKey = this.state.siteUrl.toString(); // Assuming siteUrl is a string
                            this.loadGroupNames(selectedKey);
                        });
                    }
                }}
                // errorMessage={groupName ? '' : 'This field is required.'}
            />
             {groupNameErrorMessage && <span style={{ color: 'red' }}>{groupNameErrorMessage}</span>}
        </div>
    </div>

    {/* User Name */}
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
        <div className={`${styles.formControl} customArea`}>
            <Label className='customLabel'>User Name<span style={{ color: 'red' }}> * </span></Label>
            <PeoplePicker
                context={this.props.context}
                //titleText="User Name"
                personSelectionLimit={1}
                showHiddenInUI={false}
                ensureUser={true}
                onChange={(items: any) => {
                    this.setState({ userName: items });
                }}
                defaultSelectedUsers={this.state.userName}
                showtooltip={true}
                principalTypes={[PrincipalType.User]}
                // errorMessage={userName ? '' : 'This field is required.'}
            // required={true}
            />
            {userNameErrorMessage && <span style={{ color: 'red' }}>{userNameErrorMessage}</span>}
        </div>
    </div>
    {/* Approval user name */}
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
        <div className={`${styles.formControl} customArea`}>
            <Label className='customLabel'>Approver Name<span style={{ color: 'red' }}> * </span></Label>
            <PeoplePicker
                context={this.props.context}
                //titleText="Approval User Name"
                personSelectionLimit={1}
                showHiddenInUI={false}
                ensureUser={true}
                onChange={(items: any) => {
                    this.setState({ ApprovalUser: items });
                }}
                //onChange={this.handleApproverNameChange} 
                defaultSelectedUsers={this.state.ApprovalUser}
                showtooltip={true}
                principalTypes={[PrincipalType.User]}
                // errorMessage={ApprovalUser ? '' : 'This field is required.'}
            // required={true}
            />
 {ApprovalUserErrorMessage && <span style={{ color: 'red' }}>{ApprovalUserErrorMessage}</span>}
        </div>
    </div>
    
    {/* Start Date */}
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
        <div className={`customControl ${styles.formControl}`}>
            <Label className='customLabel'>Start Date</Label>
            {/* <DatePicker
                onSelectDate={this.handleAddDateChange}
                value={this.state.addDate}                                  
            /> */}
             <DatePicker
                        // onSelectDate={this.handleAddDateChange}
                        value={this.state.addDate}
                        disabled={true}
                    />
            {/* {addDateErrorMessage && <span style={{ color: 'red' }}>{addDateErrorMessage}</span>} */}
        </div>
    </div>
   
    {/* Expire Date */}
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
        <div className={`customControl ${styles.formControl}`}>
            <Label className='customLabel'>Revoke Days<span style={{ color: 'red' }}> *</span></Label>
            {/* <DatePicker
                //label="Expire Date"
                onSelectDate={this.handleRemoveDateChange}
                value={this.state.removeDate1}
                // minDate={new Date()}
                minDate={new Date(nextdays)}
            // isRequired={true}

            //errorMessage={reason ? '' : 'This field is required.'}

            /> */}

<NumericFormat 
                          id='txtRevokeDays'
                          className='number-format'
                          
                          //disabled={this.state.isDisplayMode || !this.state.enablestatus || this.state.editDisableForDeclinePending || this.state.departmentdisbalestatus}
                          value={this.state.RevokeDays}
                          onValueChange={(e: any) =>
                            this.numberFormatTextChangeEvent("txtRevokeDays", e)
                          }
                        />
              {removeDateErrorMessage && <span style={{ color: 'red' }}>{removeDateErrorMessage}</span>}
        </div>
    </div>
    <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6"></div>
    {/* Reason */}
    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
    {/* {this.state.ApprovalUser && this.state.ApprovalUser.length > 0 && this.state.ApprovalUser[0].secondaryText !== this.props.context.pageContext.user.email ? */}
        <div className={`${styles.formControl} customArea`}>
            <Label className='customLabel'>Remarks<span style={{ color: 'red' }}> *</span> </Label>
            <TextField
                //label="Reason"
                value={this.state.reason}
                onChange={(event, newValue) => this.setState({ reason: newValue })}
                // errorMessage={this.state.isremovedate && reason ? '' : 'This field is required.'}
                multiline rows={3}
                disabled={this.state.ApprovalUser && this.state.ApprovalUser.length > 0 && this.state.ApprovalUser[0].secondaryText === this.props.context.pageContext.user.email}
            />
{reasonErrorMessage && <span style={{ color: 'red' }}>{reasonErrorMessage}</span>}
        </div>
    
    {/* : undefined } */}
    </div>

</div>


<div className={`customButton ${styles.formButton}`}>
    <PrimaryButton text="Submit" onClick={this.saveData} />
    <PrimaryButton text="Cancel" onClick={this.resetForm} />

    {this.state.isDialogVisible && (
                    <Dialog
                        hidden={!this.state.isDialogVisible}
                        onDismiss={this.dismissDialog}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Success',
                            subText : dialogSubText
                           // subText: 'Permission is assigned to this user sucessfully',
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 450 } },
                        }}
                    >
                       <DialogFooter>
            <PrimaryButton onClick={this.dismissDialog} text="OK" />
        </DialogFooter>
                    </Dialog>
                )}
</div>
</div>
</>
            
        );
    }
}
