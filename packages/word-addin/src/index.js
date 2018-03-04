import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import registerServiceWorker from './registerServiceWorker';
import { initializeIcons } from '@uifabric/icons';

// Register icons and pull the fonts from the default SharePoint cdn:
initializeIcons();

const Office = window.Office;

Office.initialize = () => {
  ReactDOM.render(<App />, document.getElementById('root'));
};

registerServiceWorker();
