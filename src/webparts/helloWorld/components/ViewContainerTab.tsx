import * as React from 'react';
import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Web} from "@pnp/sp";
import Swal from 'sweetalert2';
interface ViewContainerTabProps {
  selectedItem: any;
  siteUrl: string;
  UserName: string;
}
 
interface SecondItem {
  Title: string;
  ReferenceId: string;
  ListDescription: string;
  ListQuantity: string;
  ListAmount: string;
  ListRemark: string;
  selectedItem: any;
}
 
interface Approver {
  OrderBy: number; // Change ReactNode to number
  ApprovalLevel: string;
  names: string;
  time: string;
  notes: string;
  level: string;
  status: string;
  AssignedOn: string; // Add the AssignedOn property
}
 
interface User {
  LoginName: string;
  ApprovalLevel: string;
  // Add other properties if needed
}
 
const ViewContainerTab: React.FC<ViewContainerTabProps> = ({ selectedItem, siteUrl, UserName }) => {
  const [items, setItems] = useState<SecondItem[]>([]);
  const [approvers, setApprovers] = useState<Approver[]>([]);
  const [comment, setComment] = useState("");
  const [allowedUserLogins, setAllowedUserLogins] = useState<User[]>([]);
  const [currentUserLogin, setCurrentUserLogin] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const navigate = useNavigate();
  const [users, setUsers] = useState<User[]>([]);
  useEffect(() => {
    getCurrentUserLogin();
    getAllowedUsers();
    fetchData();
  }, [selectedItem.ReferenceId, siteUrl]);
 
  useEffect(() => {
    if (allowedUserLogins.length > 0) {
      console.log("allowedUserLogins has been updated");
    }
  }, [allowedUserLogins]);
 
  useEffect(() => {
    if (currentUserLogin !== '') {
      console.log("currentUserLogin has been updated");
    }
  }, [currentUserLogin]);
 
  useEffect(() => {
    if (users.length > 0) {
      console.log("Users have been updated");
      // Additional logic if needed...
    }
  }, [users]);
 
  async function getCurrentUserLogin() {
    try {
      const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
      console.log("Fetching current user login...");
      const user = await NewWeb.currentUser.get();
      console.log("Current user login:", user.LoginName);
 
      setCurrentUserLogin(user.LoginName);
      setIsLoading(false);
    } catch (error) {
      console.error("Error fetching current user login:", error);
      setIsLoading(false);
    }
  }
 
  const getAllowedUsers = async (): Promise<any[]> => {
    try {
      const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
      console.log("Fetching allowed users...");
      const allowedUsersList = await NewWeb.lists.getByTitle("AllowedUsers").items.get();
      setAllowedUserLogins(allowedUsersList); // Assuming allowedUsersList is an array of objects
      setIsLoading(false);
      console.log("Allowed users:", allowedUsersList);
      return allowedUsersList;
    } catch (error) {
      console.error("Error fetching allowed users:", error);
      setIsLoading(false);
      return [];
    }
  };
 
  async function fetchData() {
    try {
      const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
 
      // Fetch items from DynamicFormDetails list
      const secondItems: SecondItem[] = await NewWeb.lists.getByTitle("DynamicFormDetails")
        .items.filter(`ReferenceId eq '${selectedItem.ReferenceId}'`).getAll();
      setItems(secondItems);
 
      // Fetch items from WorkFlow list
      const itemsData = await NewWeb.lists.getByTitle("FormDetails")
        .items.filter(`ReferenceId eq '${selectedItem.ReferenceId}'`).getAll();
 
      if (itemsData.length > 0) {
        // Fetch items from WorkFlow list using the reference ID from itemsData
        const workflowData = await NewWeb.lists.getByTitle("WorkFlow")
          .items.filter(`Title eq '${itemsData[0].ReferenceId}'`).getAll();
 
        // Map workflowData to Approver interface
        const approversData: Approver[] = workflowData.map((item: any, index: number) => ({
          OrderBy: index + 1,
          names: item.names,
          time: item.time,
          notes: item.notes,
          level: item.level,
          status: item.status,
          AssignedOn: item.AssignedOn,
          ApprovalLevel: item.ApprovalLevel
        }));
        setApprovers(approversData);
      }
    } catch (error) {
      console.error('Error retrieving data:', error);
    }
  }
 
  const rejectItem = async (comment: string, currentUserLogin: string, selectedItem: any) => {
    try {
      const currentTime = new Date().toISOString();
      const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
      const [, , currentUserEmail] = currentUserLogin.split('|');
 
      const newItem = {
        status: 'rejected',
        notes: comment,
        AssignedOn: currentTime,
        Title: selectedItem?.ReferenceId,
        names: currentUserEmail,
        ApprovalLevel: 'Rejected'
      };
 
      // Add the rejected item to the "WorkFlow" list
      await NewWeb.lists.getByTitle("WorkFlow").items.add(newItem);
 
      // Get the current rejection count for the item
      const itemsData = await NewWeb.lists.getByTitle("FormDetails")
        .items.filter(`ReferenceId eq '${selectedItem.ReferenceId}'`).getAll();
      const rejectionCount = itemsData.filter(item => item.stat === 'rejected').length;
 
      // If two or more users have rejected the item, notify the next level user
      if (rejectionCount >= 2) {
        const nextLevelUser = await getCurrentUserWithLevel('2', 'level'); // Get the next level user
        if (nextLevelUser) {
          // Notify the next level user
          // Your notification code goes here
        }
      }
 
      // Assign the rejected item to the level one user
      const levelOneUser = await getCurrentUserWithLevel('1', 'level'); // Get level one user
      if (levelOneUser) {
        const levelOneItem = {
          Title: selectedItem.ReferenceId,
          names: levelOneUser.LoginName,
          ApprovalLevel: '1',
          AssignedBy: currentUserEmail,
          AssignedTo: levelOneUser.Email,
          AssignedOn: currentTime,
          status: 'rejected',
        };
        await NewWeb.lists.getByTitle("WorkFlow").items.add(levelOneItem);
      }
 
      // Update the 'stat' field of items in the 'FormDetails' list to 'rejected'
      for (const item of itemsData) {
        await NewWeb.lists.getByTitle("FormDetails").items.getById(item.Id).update({
          stat: 'rejected'
        });
      }
      fetchData();
      Swal.fire({
        icon: "success",
        title: "Your status have been rejected",
      });
 
    } catch (error) {
      console.error('Error rejecting item:', error);
    }
  };
 
  const userlevel = async () => {
    try {
      const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
      const user = await NewWeb.lists.getByTitle("AllowedUsers").items.filter(`ApprovalLevel eq '${userlevel}'`).get();
      setUsers(user);
    } catch (error) {
      console.error('Error fetching users:', error);
    }
  };
 
  const getCurrentUserWithLevel = async (value: string, searchType: 'email' | 'level'): Promise<any> => {
    try {
      console.log(`Fetching user with ${searchType === 'email' ? 'email' : 'approval level'} ${value}...`);
 
      let filterQuery = '';
      if (searchType === 'email') {
        filterQuery = `LoginName eq '${value}'`;
      } else if (searchType === 'level') {
        filterQuery = `ApprovalLevel eq '${value}'`;
      }
      const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
      const users = await NewWeb.lists.getByTitle("AllowedUsers").items.filter(filterQuery).get();
 
      if (users.length > 0) {
        return {
          ...users[0],
          ApprovalLevel: users[0].ApprovalLevel // Assuming the approval level is stored in the ApprovalLevel field
        };
      } else {
        console.error(`No user found with ${searchType === 'email' ? 'email' : 'approval level'} ${value}.`);
        return null;
      }
    } catch (error) {
      console.error(`Error fetching user with ${searchType === 'email' ? 'email' : 'approval level'} ${value}:`, error);
      return null;
    }
  };
 
 
  const getNextUsersWithLevel = async (approvalLevel: string): Promise<any[]> => {
    try {
        console.log(`Fetching users with approval level ${approvalLevel}...`);
 
        const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
        const users = await NewWeb.lists.getByTitle("AllowedUsers").items.filter(`ApprovalLevel eq '${approvalLevel}'`).get();
 
        console.log(`Fetched ${users.length} users with approval level ${approvalLevel}.`);
 
        return users;
    } catch (error) {
        console.error(`Error fetching users with approval level ${approvalLevel}:`, error);
        return [];
    }
};
 
 
const approveItem = async (comment: string, currentUserLogin: string, selectedItem: any, isResubmit: boolean = false) => {
  try {
      const currentTime = new Date().toISOString();
      const webClient = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
      const [, , currentUserEmail] = currentUserLogin.split('|');
     
      // Fetch user's level information
      const userData = await getCurrentUserWithLevel(currentUserEmail, 'email');
     
      if (!userData || !userData.ApprovalLevel) {
          throw new Error("User data or approval level is missing.");
      }
 
      const { ApprovalLevel: userLevel } = userData;
 
      // Check if all previous approvers have approved the item or it's a resubmit
      const allApproved = !isResubmit && approvers.every(approver => approver.status === 'approved');
 
      // If all previous approvers have approved the item or it's a resubmit, proceed with approval
      if (allApproved || isResubmit) {
          const newItem = {
              status: isResubmit ? 'pending' : 'approved', // Set status to pending if it's a resubmit
              notes: comment,
              AssignedOn: currentTime,
              Title: selectedItem?.ReferenceId,
              names: currentUserEmail,
              ApprovalLevel: userLevel, // Assigning the user's level dynamically
          };
 
          // Add the approved item to the "WorkFlow" list
          await webClient.lists.getByTitle("WorkFlow").items.add(newItem);
 
          // Determine the next level for assignment dynamically
          let nextLevel = '';
          if (!isResubmit) {
              // Dynamically calculate the next level in ascending order
              nextLevel = String(parseInt(userLevel) + 1);
          }
 
          // Fetch users with the next approval level dynamically
          if (nextLevel !== '') {
              const nextLevelUsers = await getNextUsersWithLevel(nextLevel);
              if (nextLevelUsers.length > 0) {
                  for (const nextLevelUser of nextLevelUsers) {
                      const nextLevelItem = {
                          Title: selectedItem.ReferenceId,
                          names: nextLevelUser.LoginName,
                          ApprovalLevel: nextLevelUser.ApprovalLevel,
                          AssignedBy: currentUserEmail,
                          AssignedTo: nextLevelUser.Email,
                          AssignedOn: currentTime,
                          status: 'pending', // Set status to pending for all users with the next level
                      };
                      await webClient.lists.getByTitle("WorkFlow").items.add(nextLevelItem);
 
                      // Notify the next level user
                      // Implement notification logic here
                  }
              }
          } else {
              // If there's no next level, update the 'status' field of items in the 'FormDetails' list to 'approved'
              // Update the UI to reflect the changes
              const NewWeb = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
              await NewWeb.lists.getByTitle("FormDetails").items.getById(selectedItem.Id).update({ stat: 'approved' });
              fetchData();
              // Call fetchData again to refresh the UI
          }
 
          // Show success message
          Swal.fire({
              icon: 'success',
              title: 'Item Approved',
              text: 'The item has been approved successfully.',
          });
 
      } else {
          // If not all previous approvers have approved or it's not a resubmit, show an error message
          console.error('Not all previous approvers have approved or it is not a resubmission.');
      }
  } catch (error) {
      console.error('Error approving item:', error);
  }
};
 
 
 
 
 
 
 
 
  const cancel = () => {
    navigate('/path/to/ViewContainerTab');
  }
  const Resubmit = async () => {
    try {
        const webClient = new Web("https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/");
 
        // Fetch the item to be deleted from the "WorkFlow" list based on Title
        const items = await webClient.lists.getByTitle("WorkFlow").items
            .filter(`Title eq '${selectedItem.ReferenceId}'`).get();
 
        if (items.length > 0) {
            // Loop through each matching item and delete it
            for (const item of items) {
                await webClient.lists.getByTitle("WorkFlow").items.getById(item.Id).delete();
            }
            fetchData();
            Swal.fire({
                icon: "success",
                title: "The item has been resubmitted successfully.",
            });
        } else {
            throw new Error("Item not found in the 'WorkFlow' list.");
        }
    } catch (error) {
        console.error('Error resubmitting item:', error);
        Swal.fire({
            icon: "error",
            title: "An error occurred while resubmitting the item.",
            text: error.message || "Unknown error",
        });
    }
};
 
 
 
 
 
  const handleCommentChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setComment(event.target.value);
  };
  if (isLoading) {
    return (
      <div>
        <img src="loader.gif" alt="Loading..." />
        <div>Loading...</div>
      </div>
    );
  }
 
  return (
    <div className="form_banner clearfix">
      <div className="header_form">
        <a href="#"><img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/next.svg" alt="Next" /></a>
        <h2>Purchase Requisition Form</h2>
      </div>
      <div className="form_block view_form">
        <div className="row">
          <div className="col-md-3">
            <div className="form-group">
              <label>Date</label>
              <p>{selectedItem.ListDate}</p>
            </div>
          </div>
          <div className="col-md-3">
            <div className="form-group">
              <label>Number</label>
              <p>{selectedItem.ListNumber}</p>
            </div>
          </div>
          <div className="col-md-3">
            <div className="form-group">
              <label>Department</label>
              <p>{selectedItem.ListDepartment}</p>
            </div>
          </div>
          <div className="col-md-3">
            <div className="form-group">
              <label>Budget</label>
              <p>{selectedItem.ListBudget}</p>
            </div>
          </div>
        </div>
        <div className="table-responsive">
          <table className="table viewform_table">
            <thead>
              <tr className="open">
                <th className="items">Item #</th>
                <th className="Description">Description <span className="required">*</span></th>
                <th className="qty">Quantity <span className="required">*</span></th>
                <th className="amnt">Amount in (AED) <span className="required">*</span></th>
                <th className="Remarks">Remarks</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, index) => (
                <tr key={index} className="open">
                  <td>{item.ReferenceId}</td>
                  <td>{item.ListDescription}</td>
                  <td>{item.ListQuantity}</td>
                  <td>{item.ListAmount}</td>
                  <td>{item.ListRemark}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div>
          <strong>Status:</strong> {selectedItem.stat}
        </div>
        <div className="app_status_block">
          <h3>Approval status</h3>
          <div>
            <table className="table viewform_table viewapp_table">
              <thead>
                <tr>
                  <th className="sno_th text-center">#</th>
                  <th className="appname_th">Approver name</th>
                  <th className="level_th">Level</th>
                  <th className="notes_th">Notes</th>
                  <th className='date_th'>Time</th>
                  <th className="status_th text-center">Status</th>
                </tr>
              </thead>
              <tbody>
                {approvers.map((approver, index) => (
                  <tr key={index}>
                    <td className="text-center">
                      <span className="dot_img"><img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/dot.svg" alt="Dot" /></span> {approver.OrderBy}
                    </td>
                    <td><p>{approver.names}</p></td>
                    <td>{approver.ApprovalLevel}</td>
                    <td>{approver.notes}</td>
                    <td>{approver.AssignedOn}</td>
                    <td className="text-center status pending">
                      <span>{approver.status}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="comments_block">
          <h4 className="comment_text">Your Comments</h4>
          <div className="comment_msg relative">
            <div className="sale_manager">Sales Manager</div>
            <textarea placeholder="Your Message" value={comment} onChange={handleCommentChange}></textarea>
            <div className="button">
              <div>
                {allowedUserLogins.some(user =>
  (user.LoginName.trim().toLowerCase() === currentUserLogin.replace('i:0#.f|membership|', '').trim().toLowerCase() &&
    (user.ApprovalLevel === '1' || user.ApprovalLevel === '2'))
) ? (
  <>
    {(approvers.every(approver => approver.status !== 'approved') || approvers.every(approver => approver.status !== 'rejected')) && (
      <button className="approve_btn" onClick={() => approveItem(comment, currentUserLogin, selectedItem)}>Approve</button>
    )}
    {(approvers.every(approver => approver.status !== 'approved') || approvers.every(approver => approver.status !== 'rejected')) && (
      <button className="rejected_btn" onClick={() => rejectItem(comment, currentUserLogin, selectedItem)}>Reject</button>
    )}
  </>
) : (
  <>
    {approvers.every(approver => approver.status !== 'approved') && approvers.every(approver => approver.status !== 'rejected') && (
      <button className="approve_btn" onClick={() => approveItem(comment, currentUserLogin, selectedItem)}>Approve</button>
    )}
    {approvers.every(approver => approver.status !== 'approved') && approvers.every(approver => approver.status !== 'rejected') && (
      <button className="rejected_btn" onClick={() => rejectItem(comment, currentUserLogin, selectedItem)}>Reject</button>
    )}
  </>
)}
{allowedUserLogins.some(user =>
  (user.LoginName.trim().toLowerCase() === currentUserLogin.replace('i:0#.f|membership|', '').trim().toLowerCase() &&
    (user.ApprovalLevel === '1' || user.ApprovalLevel === '2'))
) && (
  <button className="cancel_btn" onClick={cancel}>Cancel</button>
)}
{allowedUserLogins.some(user =>
  (user.LoginName.trim().toLowerCase() === currentUserLogin.replace('i:0#.f|membership|', '').trim().toLowerCase() &&
    (user.ApprovalLevel === '1' || user.ApprovalLevel === '2'))
) && (
  <button className="approve_btn" onClick={Resubmit}>Resubmit</button>
)}
 
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
 
export default ViewContainerTab;