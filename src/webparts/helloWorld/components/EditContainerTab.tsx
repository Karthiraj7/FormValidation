import * as React from 'react';
import { Web } from "@pnp/sp";
import Swal from 'sweetalert2';
import { useNavigate } from 'react-router-dom';

interface EditContainerTabProps {
  selectedItem: any;
  siteurl: string;
  UserName: string;
}

interface ThirdItem {
  Title: string;
  ReferenceId: string;
  ListDescription: string;
  ListQuantity: string;
  ListAmount: string;
  ListRemark: string;
  ID: number;
  ListDate: string;
  ListNumber: string;
  ListDepartment: string;
  ListBudget: string;
}

const EditContainerTab: React.FC<EditContainerTabProps> = ({ selectedItem }) => {
  const [items, setItems] = React.useState<ThirdItem[]>([]);
  const [date, setDate] = React.useState(selectedItem.ListDate);
  const [number, setNumber] = React.useState(selectedItem.ListNumber);
  const [department, setDepartment] = React.useState(selectedItem.ListDepartment);
  const [budget, setBudget] = React.useState(selectedItem.ListBudget);
  const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
  const navigate = useNavigate();
  React.useEffect(() => {
    async function getAllItems() {
      try {
        const items: ThirdItem[] = await NewWeb.lists.getByTitle("DynamicFormDetails")
          .items.filter(`ReferenceId eq '${selectedItem.ReferenceId}'`).getAll();
        setItems(items);
      } catch (error) {
        console.error('Error retrieving items:', error);
      }
    }

    getAllItems();
  }, [selectedItem]);
  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>, index: number, field: keyof ThirdItem) => {
    const newItems = [...items]; // Create a copy of the items array
    newItems[index] = { ...newItems[index], [field]: e.target.value }; 
    setItems(newItems); // Set the state with the updated array
  };
  

  const handleSave = async () => {
    try {
      const mainItemToUpdate = items[0];
  
      const mainItem = await NewWeb.lists.getByTitle("FormDetails").items
        .filter(`ReferenceId eq '${mainItemToUpdate.ReferenceId}'`)
        .get();
  
      if (mainItem.length > 0) {
        const ReferenceId = mainItem[0].ID;
  
        await NewWeb.lists.getByTitle("FormDetails").items.getById(ReferenceId).update({
          ListDate: date,
          ListNumber: number,
          ListDepartment: department,
          ListBudget: budget,
        });
  
        for (const relatedItem of items.slice(1)) {
          await NewWeb.lists.getByTitle("DynamicFormDetails").items.getById(relatedItem.ID).update({
            ListDescription: relatedItem.ListDescription,
            ListQuantity: relatedItem.ListQuantity,
            ListAmount: relatedItem.ListAmount,
            ListRemark: relatedItem.ListRemark,
          });
        }
  
        setItems([]);
        Swal.fire({
         
          icon: "success",
          title: "Your details  has been saved",
          showConfirmButton: false,
         // Show the success message for 1.5 seconds
        });
      } else {
        console.error('Main item not found with the given ReferenceId.');
        alert("Main item not found with the given ReferenceId.");
      }
      navigate('/path/to/EditContainerTab');
    } 
    catch (error) {
      console.error('Error updating items:', error);
      alert("Failed to update items. Please check the console for more details.");
    }
  };


  const handleCancel = async () => {
    Swal.fire({
      icon: "error",
      title: "Cancel",
      text: "You have cancelled the details",
    });
    navigate('/path/to/EditContainerTab');
  };
  
  return (
    <div className="form_banner clearfix">
      <div className="header_form">
        <a href="#"> <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/next.svg" /></a>
        <h2> MANUAL CREDIT NOTE REQUEST </h2>
      </div>
      <div className="form_block">
        <div className="row">
          <div className="col-md-3">
            <div className="form-group">
              <label> Date </label>
              <input type="" value={date} onChange={(e) => setDate(e.target.value)} className="form-control" />
            </div>
          </div>
          <div className="col-md-3">
            <div className="form-group">
              <label> Number vfert </label>
              <input type="text" value={number} onChange={(e) => setNumber(e.target.value)} className="form-control" placeholder="Enter Number" />
            </div>
          </div>
          <div className="col-md-3">
            <div className="form-group">
              <label> Department <span className="required" > * </span> </label>
              <input type="text" value={department} onChange={(e) => setDepartment(e.target.value)} className="form-control" placeholder="Enter Department" />
            </div>
          </div>
          <div className="col-md-3">
            <div className="form-group">
              <label> Budget <span className="required"> * </span> </label>
              <input type="text" value={budget} onChange={(e) => {
                const newBudget = parseFloat(e.target.value);
                setBudget(isNaN(newBudget) ? 0 : newBudget);
              }} className="form-control" placeholder="Enter Budget" />
            </div>
          </div>
        </div>
        <div className="table-responsive">
          <table className="table">
            <thead>
              <tr className="open">
                <th className="Description"> Description  <span className="required"> * </span> </th>
                <th className="qty"> Quantity <span className="required"> * </span> </th>
                <th className="amnt"> Amount in (AED) <span className="required"> * </span> </th>
                <th className="Remarks"> Remarks <span className="required"> * </span> </th>
                <th className="text-center"> Action </th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, index) => (
                <tr className="open" key={index}>
                  <td><input type="text" value={item.ListDescription} onChange={(e) => handleInputChange(e, index, 'ListDescription')} /></td>
                  <td><input type="text" value={item.ListQuantity} onChange={(e) => handleInputChange(e, index, 'ListQuantity')} /> </td>
                  <td><input type="text" value={item.ListAmount} onChange={(e) => handleInputChange(e, index, 'ListAmount')} /> </td>
                  <td> <input className="form-control" type="text" value={item.ListRemark} onChange={(e) => handleInputChange(e, index, 'ListRemark')} /> </td>
                  <td className="text-center">
                    <a href="#"> <img className="delete_icon" src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/delete.svg" />
                      <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/delete_img.svg" className="delete_img" />
                    </a>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="button">
          <button className="submit_btn" onClick={handleSave}> Save </button>
          <button className="cancel_btn" onClick={handleCancel} > Cancel </button>
        </div>
      </div>
    </div>
  );
}

export default EditContainerTab;
