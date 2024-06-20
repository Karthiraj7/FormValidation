import * as React from 'react';
import type { IHelloWorldProps } from './IHelloWorldProps';

export default class HeaderTab extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {

    return (
        <div>
        <div className="logo">
                <a href="#"> <img src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/logo.svg" alt="image" /> </a>
            </div>
            <div className="notification-part">
                <ul>
                    <li> <a href="#"> <img className="user_img" src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/user.svg" alt="image"/> </a> </li>
                    <li> <span> Mohammed </span> </li>
                    <li> <a href="#"> <img className="next_img" src="https://3c3tsp.sharepoint.com/sites/demosite/siteone/training11/SiteAssets/SharePoint%20Assets/img/dropdown.svg" alt="image"/> </a> </li>
                </ul>
            </div>
            </div>
    );
  }
}