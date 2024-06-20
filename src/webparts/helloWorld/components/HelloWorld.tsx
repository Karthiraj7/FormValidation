import * as React from 'react';
//import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps } from './IHelloWorldProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import ContainerTab from './ContainerTab';
import HeaderTab from './HeaderTab';
import { BrowserRouter as Router, Route, Routes, Navigate } from 'react-router-dom';
import { SPComponentLoader } from '@microsoft/sp-loader';
import EditContainerTab from './EditContainerTab';
import FormContainerTab from './FormContainerTab';
// import { Route, Router } from 'react-router-dom';

import ViewContainerTab from './ViewContainerTab'; 


SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css");
SPComponentLoader.loadCss("https://fonts.googleapis.com");
SPComponentLoader.loadCss("https://fonts.gstatic.com");
SPComponentLoader.loadCss("https://fonts.googleapis.com/css2?family=Rajdhani:wght@300;400;500;600;700&display=swap");
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
SPComponentLoader.loadCss("https://3c3tsp.sharepoint.com/sites/demosite/siteone/karthiassessment/SiteAssets/cssstyle.css");

SPComponentLoader.loadScript("https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js");
SPComponentLoader.loadScript("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js");



export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  state = {
    selectedItem: {} // Define selectedItem in the component state
  };

  public render(): React.ReactElement<IHelloWorldProps> {
    const { selectedItem } = this.state;
  
    return (
      <div>
      <div className="container clearfix">
            
           <HeaderTab siteurl={this.props.siteurl} UserName={this.props.UserName}></HeaderTab>

        </div>
   
        <div className="container">
           
         
         

           <Router>
      <Routes>
        <Route path="/containertab" element={<ContainerTab siteurl={''} UserName={''} />} />
        <Route path="/EditContainerTab" element={<EditContainerTab   selectedItem={selectedItem} siteurl={''} UserName={''}  />} />
        <Route path="/ViewContainerTab" element={<ViewContainerTab  selectedItem={selectedItem} siteUrl={''} UserName={''} />} />
        <Route path="/FormContainerTab" element={<FormContainerTab siteurl={''} UserName={''}/>} />
        <Route path="*" element={<Navigate to="/containertab" />} />
      </Routes>
    </Router>



        </div>
        </div>
    );
  }
}
