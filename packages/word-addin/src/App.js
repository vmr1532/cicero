import React, { Component } from 'react';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { List } from 'office-ui-fabric-react/lib/List';
// import { ListBasicExample } from 'List.Basic.Example';

import './App.css';
import {
  DocumentCard,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity
} from 'office-ui-fabric-react/lib/DocumentCard';

class App extends Component {
  constructor(props) {
    super(props);

    this.onSetColor = this.onSetColor.bind(this);
  }

  onSetColor() {
    window.Word.run(async (context) => {
      await context.sync();
    });
  }
  
  render() {
    return (
      <div>
      <DocumentCard onClickHref='http://bing.com'>
        <DocumentCardPreview
          previewImages={ [
            {
              previewImageSrc: require('./documentpreview.png'),
              iconSrc: require('./iconppt.png'),
              width: 318,
              height: 196,
              accentColor: '#ce4b1f'
            }
          ] }
        />
        <DocumentCardTitle title='Revenue stream proposal fiscal year 2016 version02.pptx'/>
        <DocumentCardActivity
          activity='Created Feb 23, 2016'
          people={
            [
              { name: 'Kat Larrson', profileImageSrc: require('./avatarkat.png') }
            ]
          }
          />
      </DocumentCard>
      </div>
    );
  }
}

export default App;