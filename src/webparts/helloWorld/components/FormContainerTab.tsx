import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { sp } from "@pnp/sp";
import Swal from 'sweetalert2';
import * as React from 'react';


interface IContainerState {
  isPurchaseRequisitionClicked: boolean;
  additionalRows: number;
  AllItems: string;
  FormDate: string;
  FormDescription: string;
  FormDepartment: string;
  FormBudget: string;
  FormQuantity: string;
  FormAmount: string;
  FormRemarks: string;
  FormNumber: string;
  submittedData: IUserFormData[];
}

interface IUserFormData {
  ListDate: string;
  ListDescription: string;
  ListDepartment: string;
  ListBudget: string;
  ListQuantity: string;
  ListAmount: string;
  ListRemarks: string;
  ListNumber: string;
}

interface IHelloWorldProps {
  siteurl: string;
  UserName: string;
}

const FormContainerTab: React.FC<IHelloWorldProps> = (props: IHelloWorldProps) => {
  const [state, setState] = useState<IContainerState>({
    isPurchaseRequisitionClicked: false ,
    additionalRows: 1,
    AllItems: '',
    FormDate: '',
    FormDescription: '',
    FormDepartment: '',
    FormBudget: '',
    FormQuantity: '',
    FormAmount: '',
    FormRemarks: '',
    FormNumber: '',
    submittedData: [
      {
        ListDate: '',
        ListDescription: '',
        ListDepartment: '',
        ListBudget: '',
        ListQuantity: '',
        ListAmount: '',
        ListRemarks: '',
        ListNumber: '',
      },
    ],
  });
  // const [submitted, setSubmitted] = useState<boolean>(true);
  const navigate = useNavigate();
  // const [submitteds, setSubmitteds] = useState<boolean>(false);
  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setState((prevState) => ({
      ...prevState,
      [name]: value,
    }));
  };

const handleCancelClick = async() =>{
  Swal.fire({
    icon: "error",
    title: "Cancel",
    text: "You have cancelled the details",
  });
  navigate('/path/to/FormContainerTab');
}



  const handleSubmitClick = async () => {
    const { FormDate, FormDescription, FormDepartment, FormBudget, FormNumber, additionalRows, submittedData,FormQuantity,FormRemarks,FormAmount } = state;
    let referenceCount = parseInt(localStorage.getItem('lastReferenceCount') || '0', 10);

    const generateReferenceId = (): string => {
      // Increment reference count
      referenceCount++;
      // Store the updated reference count in local storage
      localStorage.setItem('lastReferenceCount', String(referenceCount));
      
      // Generate reference ID with padded count
      const paddedCount = String(referenceCount).padStart(7, '0');
      const referenceId = `tmax${paddedCount}`;
      return referenceId;
    };
    
    try {
      const referenceId = generateReferenceId(); // Generate reference ID
    if (!referenceId) {
      console.error("Error: ReferenceId could not be generated.");
      return;
    }
      // Add FormDetails item
      const formDetailsItem = {
        ListDate: FormDate,
        ListNumber: FormNumber,
        ListDepartment: FormDepartment,
        ListBudget: FormBudget,
        ReferenceId: referenceId,
      };
      const formDetailsResponse = await sp.web.lists.getByTitle("FormDetails").items.add(formDetailsItem);
      const formDetailsAddedItemId = formDetailsResponse.data.Id;
      console.log('FormDetails New item ID:', formDetailsAddedItemId);
  
      // Add DynamicFormDetails item for the main row
      const dynamicFormDetailsItem = {
        ListDescription: FormDescription,
        ListQuantity: FormQuantity,
        ListAmount: FormAmount,
        ListRemarks: FormRemarks,
        ReferenceId: referenceId,
      };
      const dynamicFormDetailsResponse = await sp.web.lists.getByTitle("DynamicFormDetails").items.add(dynamicFormDetailsItem);
      const dynamicFormDetailsAddedItemId = dynamicFormDetailsResponse.data.Id;
      console.log('DynamicFormDetails New item ID:', dynamicFormDetailsAddedItemId);
  
      // Add DynamicFormDetails items for additional rows
      for (let i = 0; i < additionalRows; i++) {
        const additionalRowItem = {
          ListDescription: submittedData[i].ListDescription,
          ListQuantity: submittedData[i].ListQuantity,
          ListAmount: submittedData[i].ListAmount,
          ListRemarks: submittedData[i].ListRemarks,
          ReferenceId: referenceId,
        };
        const additionalRowResponse = await sp.web.lists.getByTitle("DynamicFormDetails").items.add(additionalRowItem);
        console.log('AdditionalRow New item ID:', additionalRowResponse.data.Id);
      }
  
      // Reset the form state to its initial state, but keep the first row
      setState({
        isPurchaseRequisitionClicked: false,
        additionalRows: 1,
        AllItems: '',
        FormDate: '',
        FormDescription: '',
        FormDepartment: '',
        FormBudget: '',
        FormQuantity: '',
        FormAmount: '',
        FormRemarks: '',
        FormNumber: '',
        submittedData: [
          {
            ListDate: '',
            ListDescription: '',
            ListDepartment: '',
            ListBudget: '',
            ListQuantity: '',
            ListAmount: '',
            ListRemarks: '',
            ListNumber: '',
          },
        ],
      });
      // setSubmitted(false);
      // setSubmitteds(true);
      // Show success message
      Swal.fire("Good job!", "Form submitted successfully!", "success");
      // Navigate to edit page
      navigate('/path/to/FormContainerTab');
      console.log('Submitted successfully to SharePoint!');
    } catch (error) {
      console.error("Error Submitting to SharePoint:", error);
      // Handle error
    }
  };

  const handleRowChange = (e: React.ChangeEvent<HTMLInputElement>, index: number, fieldName: string) => {
    const { value } = e.target;
    setState((prevState) => {
      const newSubmittedData = prevState.submittedData.map((item, i) => {
        if (i !== index) return item;
        return { ...item, [fieldName]: value };
      });
      return { ...prevState, submittedData: newSubmittedData };
    });
  };

  const handleDeleteRow = (event: React.MouseEvent<HTMLAnchorElement>) => {
    event.preventDefault();
    event.stopPropagation();
    const index = parseInt(event.currentTarget.getAttribute('data-index') || '0', 10);
    setState((prevState) => {
      const newData = prevState.submittedData.filter((_, i) => i !== index);
      return {
        ...prevState,
        submittedData: newData,
        additionalRows: newData.length,
      };
    });
  };

  const handleAddNew = () => {
    setState((prevState) => ({
      ...prevState,
      additionalRows: prevState.additionalRows + 1,
      submittedData: [
        ...prevState.submittedData,
        {
          ListDate: '',
          ListDescription: '',
          ListDepartment: '',
          ListBudget: '',
          ListQuantity: '',
          ListAmount: '',
          ListRemarks: '',
          ListNumber: '',
        },
      ],
    }));
  };

  const renderAdditionalRows = () => {
    const { submittedData } = state;
    return submittedData.map((rowData, index) => (
      <tr key={index}>
        <td> {index + 1} </td>
        <td>
          <input
            className="form-control"
            type="text"
            placeholder="Enter Description"
            name={`additionalRowDescription${index}`}
            value={rowData.ListDescription || ''}
            onChange={(e) => handleRowChange(e, index, 'ListDescription')}
          />
        </td>
        <td>
          <input
            className="form-control"
            type="text"
            placeholder="Quantity"
            name={`additionalRowQuantity${index}`}
            value={rowData.ListQuantity || ''}
            onChange={(e) => handleRowChange(e, index, 'ListQuantity')}
          />
        </td>
        <td>
          <input
            className="form-control"
            type="text"
            placeholder="Amount"
            name={`additionalRowAmount${index}`}
            value={rowData.ListAmount || ''}
            onChange={(e) => handleRowChange(e, index, 'ListAmount')}
          />
        </td>
        <td>
          <input
            className="form-control"
            type="text"
            placeholder="Remarks"
            name={`additionalRowRemarks${index}`}
            value={rowData.ListRemarks || ''}
            onChange={(e) => handleRowChange(e, index, 'ListRemarks')}
          />
        </td>
        <td className="text-center">
          {index === 0 ? (
            <a href="#">
              <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/delete.svg" alt="Delete" />
            </a>
          ) : (
            <a href="#" data-index={index} onClick={handleDeleteRow}>
              <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/delete_img.svg" alt="Delete" />
            </a>
          )}
        </td>
      </tr>
    ));
  };

  return (
    <div>
 
{/* 
      {!submitted && ( */}
        <div id="formbannerclearfix">
          <>
            <div className="header_form">
              <a href="#"> <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/next.svg" alt="Next" /></a>
              <h2> MANUAL CREDIT NOTE REQUEST </h2>
            </div>
            <div className="form_block">
              <div className="row">
                <div className="col-md-3">
                  <div className="form-group">
                    <label> Date </label>
                    <input type="date" value={state.FormDate} name="FormDate" onChange={handleChange} className="form-control" />
                  </div>
                </div>
                <div className="col-md-3">
                  <div className="form-group">
                    <label> Number </label>
                    <input type="text" value={state.FormNumber} name="FormNumber" onChange={handleChange} className="form-control" placeholder="Enter Number" />
                  </div>
                </div>
                <div className="col-md-3">
                  <div className="form-group">
                    <label> Department </label>
                    <input type="text" value={state.FormDepartment} name="FormDepartment" onChange={handleChange} className="form-control" placeholder="Enter Department" />
                  </div>
                </div>
                <div className="col-md-3">
                  <div className="form-group">
                    <label> Budget </label>
                    <input type="text" value={state.FormBudget} name="FormBudget" onChange={handleChange} className="form-control" placeholder="Enter Budget" />
                  </div>
                </div>
              </div>
              <div className="table-responsive">
                <table className="table">
                  <thead>
                    <tr className="open">
                      <th className="items"> Item # </th>
                      <th className="Description"> Description <span className="required"> * </span> </th>
                      <th className="qty"> Quantity <span className="required"> * </span> </th>
                      <th className="amnt"> Amount in (AED) <span className="required"> * </span> </th>
                      <th className="Remarks"> Remarks </th>
                      <th className="text-center"> Action </th>
                    </tr>
                  </thead>
                  <tbody>
                    {renderAdditionalRows()}
                  </tbody>
                </table>
              </div>
              <div className="Add_new"> <a href="#" onClick={handleAddNew}> Add New </a></div>
              <div className="button">
                <button className="submit_btn" onClick={handleSubmitClick}> Submit </button>
                <button className="cancel_btn"  onClick={handleCancelClick} > Cancel </button>
              </div>
            </div>
          </>
        </div>
      {/* )} */}
    </div>
  );
};

export default FormContainerTab;
