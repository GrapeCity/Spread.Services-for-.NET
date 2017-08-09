import * as React from 'react';
import { Link, NavLink } from 'react-router-dom';

export class Home extends React.Component<{}, {}> {
    public render() {
        return <div>
            <h1>GrapeCity Spread.Services Asp.Net Core Demo!</h1>
            <p>Welcome to Spread.Services Asp.net Core demo, you can learn how to progam with Spread.Services in Asp.Net Core at server side, and use React + Spread.Sheets at client side:</p>
            <ul>
                <li><a href='https://www.nuget.org/packages/Spread.Services/1.0.0-rc'>Spread.Services</a> is a new high performing, low memory server component with full API for server-side spreadsheet generation, manipulation, and serialization to various formats including xlsx and ssjson. Spread.Services targets .NET Standard 1.4 for multi-platform support including: .NET Framework, .NET Core, and Mono.</li>
                <li><a href='http://spread.grapecity.com/Demos/JS/TutorialSample/'>Spread.Sheets</a> is the spreadsheet component of the SpreadJS product family. This enterprise-grade JavaScript spreadsheet displays and manages data much like Microsoft Excel. Popular features include a formula engine, sorting, filtering, input controls, sparklines, and native Excel input/output.</li>
            </ul>
            <p>You will explore some typical seneros about how to use Spread.Services together with Spread.Sheets:</p>
            <ul>
                <li>In <Link to={'/ExcelTemplateDemo'}>Excel Template Demo</Link>, you will learn how to open an excel template at server side with Spread.Services, then view or fill content for the template through Spread.Sheets at client side.</li>
                <li>In <Link to={'/ProgrammingDemo'}>Programming API Demo</Link>, You will learn how to program with Spread.Services all yourself at server side, then view the result through Spread.Sheets at client side.</li>
                <li>In <Link to={'/ExcelIODemo'}>Excel IO Demo</Link>, you will learn how to upload an excel file from client and open the file using Spread.Services at server side, then view the result through Spread.Sheets at client side.</li>
            </ul>
            <p>You can find more resources about <strong>Spread.Services</strong> at:</p>
            <ul>
                <li><a href='http://spread.cloudapp.net/spreadservices'>Online Demo Site</a></li>
                <li><a href='http://spread.grapecity.com'>Product Home Site</a></li>
                <li><a href='https://nuget.org/packages/Spread.Services/'>Spread.Services Nuget Package Site</a></li>
                <li><a href='http://spread.grapecity.com/spreadjs/sheets/'>Spread.Sheets Home Site</a></li>
            </ul>
        </div>;
    }
}
