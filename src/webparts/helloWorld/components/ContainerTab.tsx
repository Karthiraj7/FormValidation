import { useState, useEffect } from 'react';
import * as React from 'react';
import type { IHelloWorldProps } from './IHelloWorldProps';
import { ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { sp } from "@pnp/sp";
import SearchIcon from '@mui/icons-material/Search';
import Pagination from '@mui/material/Pagination';
import Stack from '@mui/material/Stack';
import EditIcon from "@mui/icons-material/Edit"
// import { useNavigate } from 'react-router-dom';
import Swal from 'sweetalert2';

import ViewContainerTab from './ViewContainerTab';
import EditContainerTab from './EditContainerTab';
import FormContainerTab from './FormContainerTab';

interface ListItem {
  ListDate: string;
  ListNumber: string;
  ListDepartment: string;
  ListBudget: number;
  stat?: string;
  ReferenceId: string;
}
interface User {
  LoginName: string;
  ApprovalLevel: string;
}
const ContainerTab: React.FC<IHelloWorldProps> = (props: IHelloWorldProps) => {

  // const navigate = useNavigate();
  const [items, setItems] = useState<ListItem[]>([]);
  const [filteredItems, setFilteredItems] = useState<ListItem[]>([]);
  const [currentPage, setCurrentPage] = useState<number>(0);
  const [itemsPerPage] = useState<number>(5);
  const [selectedOption, setSelectedOption] = useState<string>('');
  const [completeCount, setCompleteCount] = useState<number>(0);
  const [pendingCount, setPendingCount] = useState<number>(0);
  const [totalCount, setTotalCount] = useState<number>(0);
  const [selectedItem, setSelectedItem] = useState<any>(null);
  const [showEditContainerTab, setShowEditContainerTab] = useState<boolean>(false); 
  const [showViewContainerTab, setViewContainerTab] = useState<boolean>(false); 
  const [containerTabsVisible, setContainerTabsVisible] = useState (true);
  const [showFormContainerTab, setShowFormContainerTab] = useState<boolean>(false); 
  const [allowedUserLogins, setAllowedUserLogins] = useState<User[]>([]);
  const [currentUserLogin, setCurrentUserLogin] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);


   useEffect(() => {
    getItemsFromList();
    getCurrentUserLogin();
    getAllowedUsers();
  }, []);

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

  async function getCurrentUserLogin() {
    try {
      console.log("Fetching current user login...");
      const user = await sp.web.currentUser.get();
      console.log("Current user login:", user.LoginName,user.ApprovalLevel);
      setCurrentUserLogin(user.LoginName);
      setIsLoading(false);
    } catch (error) {
      console.error("Error fetching current user login:", error);
      setIsLoading(false);
    }
  }

  async function getAllowedUsers() {
    try {
        console.log("Fetching allowed users...");
        const allowedUsersList = await sp.web.lists.getByTitle("AllowedUsers").items.get();
        setAllowedUserLogins(allowedUsersList); // Assuming allowedUsersList is an array of objects
        setIsLoading(false);
        console.log("Allowed users:", allowedUsersList);
    } catch (error) {
        console.error("Error fetching allowed users:", error);
        setIsLoading(false);
    }
}




  const getItemsFromList = async () => {
    try {
      setIsLoading(true);
      console.log("Fetching items from list...");
      sp.setup({
        sp: {
          baseUrl: "https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        },
      });

      const list = sp.web.lists.getByTitle("FormDetails");
      const items = await list.items.get();
      console.log("Items fetched:", items);

      let completeCount = 0;
      let pendingCount = 0;

      const formattedItems: ListItem[] = await Promise.all(items.map(async (item: any) => {
        if (!item.stat || item.stat === "pending") {
          item.stat = 'pending';
          pendingCount++;
          await list.items.getById(item.Id).update({ stat: 'pending' });
        } else if (item.stat === "complete") {
          completeCount++;
        } else if (item.stat === "pending") {
          pendingCount++;
        }

        return {
          ListDate: new Date(item.ListDate).toLocaleDateString(),
          ListNumber: item.ListNumber,
          ListDepartment: item.ListDepartment,
          ListBudget: item.ListBudget,
          stat: item.stat,
          ReferenceId: item.ReferenceId || '',
        };
      }));

      const totalCount = items.length;

      setItems(formattedItems);
      setFilteredItems(formattedItems);
      setCompleteCount(completeCount);
      setPendingCount(pendingCount);
      setTotalCount(totalCount);
    } catch (error) {
      console.error("Error retrieving items from list:", error);
    }
  };

  const exportToExcel = () => {
    let itemsToExport = filteredItems;
    if (selectedOption === 'Complete' || selectedOption === 'Pending') {
      itemsToExport = filteredItems.filter((item: ListItem) => item.stat === selectedOption.toLowerCase());
    }

    const worksheet = XLSX.utils.json_to_sheet(itemsToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "List Items");
    XLSX.writeFile(workbook, "list_items.xlsx");
  };

  const handlePageChange = (event: ChangeEvent<unknown>, newPage: number) => {
    setCurrentPage(newPage - 1);
  };

  const handleOptionChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const selectedOption = event.target.value;
    setSelectedOption(selectedOption);
    filterAndSortItems(selectedOption);
  };

  const filterAndSortItems = (selectedOption: string) => {
    let filteredItems = items.slice();

    if (selectedOption === 'Complete' || selectedOption === 'Pending') {
      filteredItems = filteredItems.filter((item: ListItem) =>
        item.stat?.toLowerCase().includes(selectedOption.toLowerCase())
      );
    }

    if (selectedOption === 'Pending' || selectedOption === 'Complete') {
      filteredItems.sort((a: ListItem, b: ListItem) => {
        return (a.stat || '').localeCompare(b.stat || '');
      });
    }

    setFilteredItems(filteredItems);
  };

  const handleSearch = (event: ChangeEvent<HTMLInputElement>) => {
    const searchKeyword = event.target.value.toLowerCase();
    const filteredItems = items.filter((item: ListItem) =>
      item.ListDepartment.toLowerCase().includes(searchKeyword) ||
      item.ReferenceId.toLowerCase().includes(searchKeyword) ||  
      item.ListNumber.toLowerCase().includes(searchKeyword)
    );
    setFilteredItems(filteredItems);
  };
  
 
 

 
  
  const handlePurchaseRequisitionClick = () => {
    Swal.fire({
      icon: "success",
      title: "You Have Clicked Purchase Requisition",
      showConfirmButton: false,
    }).then(() => {
      setContainerTabsVisible(false);
      setShowFormContainerTab(true);
    });


   
    // navigate('/FormContainerTab');
  };
  
  

  const handleEditClick = (item: any) => {
    setSelectedItem(item);
    
    setContainerTabsVisible(false);
    setShowEditContainerTab(true);

    Swal.fire({
        title: "Sweet!",
        text: "YOU HAVE CLICKED EDIT FORM",
        imageUrl: "https://unsplash.it/400/200",
        imageWidth: 400,
        imageHeight: 200,
        imageAlt: "Custom image"
    });
};


  
const handleActionClick = (item: any) => {
  Swal.fire({
      title: "You have clicked view form",
      showClass: {
          popup: `
              animate__animated
              animate__fadeInUp
              animate__faster
          `
      },
      hideClass: {
          popup: `
              animate__animated
              animate__fadeOutDown
              animate__faster
          `
      },
      imageUrl: "https://media.tenor.com/cxsA-a-8uz0AAAAC/tom-and-jerry-jerry-the-mouse.gif",
      imageWidth: 200, // Adjust as needed
      imageHeight: 200, // Adjust as needed
  }).then(() => {
      setSelectedItem(item);
      setViewContainerTab(true);
      setContainerTabsVisible(false);
  });
};

  

  const indexOfLastItem = (currentPage + 1) * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const adjustedIndexOfLastItem = Math.min(indexOfLastItem, filteredItems.length);
  const adjustedIndexOfFirstItem = Math.min(indexOfFirstItem, filteredItems.length);
  const currentItems = filteredItems.slice(adjustedIndexOfFirstItem, adjustedIndexOfLastItem);
  if (isLoading) {
    return (
      <div>
        <img src="loader.gif" alt="Loading..." />
        <div>Loading...</div>
      </div>
    );
  }
  
  return (
    <>
      {showFormContainerTab && <FormContainerTab siteurl={''} UserName={''} />}
      {showEditContainerTab && <EditContainerTab selectedItem={selectedItem} siteurl={''} UserName={''} />}
      {showViewContainerTab && <ViewContainerTab selectedItem={selectedItem} siteUrl={''} UserName={''} />}
      {containerTabsVisible && (
        <div className="dashboard-wrap">
          <div className="heading-block clearfix">
            <h2>Dashboard</h2>
            <a href="#" className="purchase_btn" onClick={handlePurchaseRequisitionClick}>Purchase Requisition</a>
          </div>
          <div className="three-blocks-wrap">
            <div className="row">
              <div className="col-md-4">
                <div className="three-blocks">
                  <div className="three-blocks-img">
                    <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/Approved.svg" alt="image" />
                  </div>
                  <div className="three-blocks-desc">
                    <h3>{totalCount}</h3>
                    <p>Total</p>
                  </div>
                </div>
              </div>
              <div className="col-md-4">
                <div className="three-blocks">
                  <div className="three-blocks-img">
                    <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/pending.svg" alt="image" />
                  </div>
                  <div className="three-blocks-desc">
                    <h3>{pendingCount}</h3>
                    <p>Total Pending</p>
                  </div>
                </div>
              </div>
              <div className="col-md-4">
                <div className="three-blocks">
                  <div className="three-blocks-img">
                    <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/rejected.svg" alt="image" />
                  </div>
                  <div className="three-blocks-desc">
                    <h3>{completeCount}</h3>
                    <p>Total Complete</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="table-wrap">
          <div className="table-search-wrap clearfix">
                <div className="table-search relative">
                <SearchIcon />
                <input type="text" onChange={handleSearch} placeholder="Search" className="" />
      
                </div>
              <div className="table-sort">
                <ul>
                  <li>
                    <span>Export to</span>
                    <a href="#" onClick={exportToExcel}><img className="excel_img" src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/excel.svg" alt="excel icon" /></a>
                  </li>
                  <li>
                    <span>Sort By</span>
                    <select name="sort_by" id="sort_by" onChange={handleOptionChange}>
                      <option value=""></option>
                      <option value="Complete">Complete</option>
                      <option value="Pending">Pending</option>
                    </select>
                  </li>
                </ul>
              </div>
            </div>
            <div className="table-responsive">
              <table className="table dashboard_table">
                <thead>
                  <tr>
                    <th className="s_no">ReferanceID</th>
                    <th className="nuber">ListDate</th>
                    <th className="date">ListNumber</th>
                    <th className="dept">ListDepartment</th>
                 
                    <th className="Purpose">Status</th>
                    <th className="Items text-center">View</th>
                    <th className="text-center status">Edit</th>
                  </tr>
                </thead>
                <tbody>
                  {currentItems.map((item, index) => (
                    <tr key={index}>
                      <td>{item.ReferenceId}</td>
                      <td>{new Date(item.ListDate).toLocaleDateString()}</td>
                      <td>{item.ListNumber}</td>
                      <td>{item.ListDepartment}</td>
                 
                      <td>{item.stat}</td>
                      <td>
                        {allowedUserLogins.some(user =>
                          user.LoginName.trim().toLowerCase() === currentUserLogin.replace('i:0#.f|membership|', '').trim().toLowerCase() &&
                          user.ApprovalLevel === '1' || user.ApprovalLevel === '2'
                        ) ? (
                          <a href="#" onClick={() => handleActionClick(item)}>
                            <img className="status approved text-center" src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/view.svg" alt="view icon" />
                          </a>
                        ) : (
                          <EditIcon className="text-center" onClick={() => handleEditClick(item)} />
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <Stack spacing={2} sx={{ alignItems: 'center' }}>
              <Pagination
                count={Math.ceil(filteredItems.length / itemsPerPage)}
                page={currentPage}
                onChange={handlePageChange}
              />
            </Stack>
          </div>
        </div>
      )}
    </>
  );
  
  
};

export default ContainerTab;
