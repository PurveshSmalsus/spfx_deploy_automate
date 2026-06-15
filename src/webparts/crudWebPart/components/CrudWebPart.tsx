import * as React from 'react';
import styles from './CrudWebPart.module.scss';
import { ICrudWebPartProps } from './ICrudWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IListItem {
  Id?: number;
  Title: string;
  Description?: string;
}

export interface ICrudWebPartState {
  items: IListItem[];
  listExists: boolean;
  loading: boolean;
  statusMessage: string;
  isEditing: boolean;
  currentItem: IListItem;
  showCreateListButton: boolean;
}

export default class CrudWebPart extends React.Component<ICrudWebPartProps, ICrudWebPartState> {
  private readonly listName = 'SPFxCRUDList';

  constructor(props: ICrudWebPartProps) {
    super(props);
    this.state = {
      items: [],
      listExists: false,
      loading: true,
      statusMessage: '',
      isEditing: false,
      currentItem: { Title: '', Description: '' },
      showCreateListButton: false
    };
  }

  public componentDidMount(): void {
    this._checkListExistence();
  }

  // Check if the SharePoint List exists
  private _checkListExistence(): void {
    this.setState({ loading: true, statusMessage: 'Checking SharePoint list...' });
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')`;

    this.props.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          this.setState({ listExists: true, showCreateListButton: false }, () => {
            this._fetchItems();
          });
        } else if (response.status === 404) {
          this.setState({
            listExists: false,
            loading: false,
            showCreateListButton: true,
            statusMessage: `List '${this.listName}' not found. Please click below to create it.`
          });
        } else {
          this.setState({
            loading: false,
            statusMessage: `Error checking list: ${response.statusText}`
          });
        }
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error checking list: ${error.message}`
        });
      });
  }

  // Create SharePoint List programmatically with 'Description' field
  private _createList = (): void => {
    this.setState({ loading: true, statusMessage: `Creating SharePoint list '${this.listName}'...` });
    const url = `${this.props.siteUrl}/_api/web/lists`;

    const listData = {
      AllowContentTypes: true,
      BaseTemplate: 100,
      Description: 'List for SPFx CRUD Web Part demo.',
      Title: this.listName
    };

    this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(listData)
    })
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          // List created, now add custom 'Description' field
          this._createDescriptionField();
        } else {
          this.setState({
            loading: false,
            statusMessage: `Failed to create list: ${response.statusText}`
          });
        }
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error creating list: ${error.message}`
        });
      });
  }

  // Add custom 'Description' field to the list
  private _createDescriptionField(): void {
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/fields`;
    const fieldData = {
      Title: 'Description',
      FieldTypeKind: 3, // Multiple lines of text
      Required: false
    };

    this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(fieldData)
    })
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          this.setState({
            listExists: true,
            showCreateListButton: false,
            statusMessage: `List '${this.listName}' created successfully!`
          }, () => {
            this._fetchItems();
          });
        } else {
          this.setState({
            loading: false,
            statusMessage: `List created, but failed to create Description field: ${response.statusText}`
          });
        }
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error creating Description field: ${error.message}`
        });
      });
  }

  // Fetch all items from SharePoint List
  private _fetchItems(): void {
    this.setState({ loading: true, statusMessage: 'Loading items...' });
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items?$select=Id,Title,Description`;

    this.props.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          return response.json();
        } else {
          throw new Error(`Failed to load items: ${response.statusText}`);
        }
      })
      .then(data => {
        this.setState({
          items: data.value,
          loading: false,
          statusMessage: data.value.length === 0 ? 'No items found. Add some!' : ''
        });
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error loading items: ${error.message}`
        });
      });
  }

  // Handle Input Changes
  private _onInputChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
    const { name, value } = event.target;
    this.setState(prevState => ({
      currentItem: {
        ...prevState.currentItem,
        [name]: value
      }
    }));
  }

  // Create an Item
  private _createItem = (event: React.FormEvent): void => {
    event.preventDefault();
    const { currentItem } = this.state;
    if (!currentItem.Title.trim()) return;

    this.setState({ loading: true, statusMessage: 'Adding item...' });
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items`;

    this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify({
        Title: currentItem.Title,
        Description: currentItem.Description
      })
    })
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          this.setState({
            currentItem: { Title: '', Description: '' },
            statusMessage: 'Item added successfully!'
          }, () => {
            this._fetchItems();
          });
        } else {
          this.setState({
            loading: false,
            statusMessage: `Failed to add item: ${response.statusText}`
          });
        }
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error adding item: ${error.message}`
        });
      });
  }

  // Edit Item (loads to state)
  private _editItem = (item: IListItem): void => {
    this.setState({
      isEditing: true,
      currentItem: { ...item }
    });
  }

  // Cancel Editing
  private _cancelEdit = (): void => {
    this.setState({
      isEditing: false,
      currentItem: { Title: '', Description: '' }
    });
  }

  // Update Item
  private _updateItem = (event: React.FormEvent): void => {
    event.preventDefault();
    const { currentItem } = this.state;
    if (!currentItem.Id || !currentItem.Title.trim()) return;

    this.setState({ loading: true, statusMessage: 'Updating item...' });
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items(${currentItem.Id})`;

    this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: currentItem.Title,
        Description: currentItem.Description
      })
    })
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          this.setState({
            currentItem: { Title: '', Description: '' },
            isEditing: false,
            statusMessage: 'Item updated successfully!'
          }, () => {
            this._fetchItems();
          });
        } else {
          this.setState({
            loading: false,
            statusMessage: `Failed to update item: ${response.statusText}`
          });
        }
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error updating item: ${error.message}`
        });
      });
  }

  // Delete Item
  private _deleteItem = (id: number): void => {
    if (!confirm('Are you sure you want to delete this item?')) return;

    this.setState({ loading: true, statusMessage: 'Deleting item...' });
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items(${id})`;

    this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    })
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          this.setState({
            statusMessage: 'Item deleted successfully!'
          }, () => {
            this._fetchItems();
          });
        } else {
          this.setState({
            loading: false,
            statusMessage: `Failed to delete item: ${response.statusText}`
          });
        }
      })
      .catch(error => {
        this.setState({
          loading: false,
          statusMessage: `Error deleting item: ${error.message}`
        });
      });
  }

  public render(): React.ReactElement<ICrudWebPartProps> {
    const {
      isDarkTheme,
      hasTeamsContext
    } = this.props;

    const {
      items,
      listExists,
      loading,
      statusMessage,
      isEditing,
      currentItem,
      showCreateListButton
    } = this.state;

    return (
      <section className={`${styles.crudWebPart} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.darkTheme : ''}`}>
        <div className={styles.container}>
          
          {/* Header Card */}
          <div className={styles.header}>
            <h1>SharePoint List CRUD Operations</h1>
            <p>Interacting with List: <strong>{this.listName}</strong></p>
            {statusMessage && (
              <div className={`${styles.statusBadge} ${loading ? styles.badgeInfo : styles.badgeSuccess}`}>
                {statusMessage}
              </div>
            )}
          </div>

          {/* Setup App Catalog and SharePoint List */}
          {showCreateListButton && (
            <div className={styles.setupCard}>
              <h3>SharePoint List Required</h3>
              <p>The web part needs a SharePoint custom list named <strong>{this.listName}</strong> to perform CRUD actions.</p>
              <button 
                className={styles.primaryButton}
                onClick={this._createList}
                disabled={loading}
              >
                {loading ? 'Creating...' : 'Auto-Create List & Fields'}
              </button>
            </div>
          )}

          {/* Form & Items Grid */}
          {listExists && (
            <div className={styles.appGrid}>
              
              {/* Form Card */}
              <div className={styles.card}>
                <h2>{isEditing ? 'Edit Item' : 'Add New Item'}</h2>
                <form onSubmit={isEditing ? this._updateItem : this._createItem} className={styles.form}>
                  <div className={styles.inputGroup}>
                    <label htmlFor="itemTitle">Title</label>
                    <input
                      type="text"
                      id="itemTitle"
                      name="Title"
                      value={currentItem.Title}
                      onChange={this._onInputChange}
                      placeholder="Enter title..."
                      required
                      disabled={loading}
                    />
                  </div>

                  <div className={styles.inputGroup}>
                    <label htmlFor="itemDescription">Description</label>
                    <textarea
                      id="itemDescription"
                      name="Description"
                      value={currentItem.Description || ''}
                      onChange={this._onInputChange}
                      placeholder="Enter description..."
                      rows={4}
                      disabled={loading}
                    />
                  </div>

                  <div className={styles.buttonGroup}>
                    <button 
                      type="submit" 
                      className={styles.primaryButton}
                      disabled={loading || !currentItem.Title.trim()}
                    >
                      {isEditing ? 'Save Changes' : 'Create Item'}
                    </button>
                    {isEditing && (
                      <button 
                        type="button" 
                        className={styles.secondaryButton} 
                        onClick={this._cancelEdit}
                        disabled={loading}
                      >
                        Cancel
                      </button>
                    )}
                  </div>
                </form>
              </div>

              {/* Items List Card */}
              <div className={styles.card}>
                <h2>Items List ({items.length})</h2>
                {loading && items.length === 0 ? (
                  <div className={styles.spinnerContainer}>
                    <div className={styles.spinner}></div>
                    <p>Loading items...</p>
                  </div>
                ) : items.length === 0 ? (
                  <div className={styles.emptyState}>
                    <p>No items found in <strong>{this.listName}</strong>.</p>
                  </div>
                ) : (
                  <div className={styles.itemsList}>
                    {items.map(item => (
                      <div key={item.Id} className={styles.itemRow}>
                        <div className={styles.itemDetails}>
                          <h3>{item.Title}</h3>
                          <p>{item.Description || <em>No description provided</em>}</p>
                        </div>
                        <div className={styles.itemActions}>
                          <button 
                            className={styles.iconButtonEdit}
                            onClick={() => this._editItem(item)}
                            title="Edit Item"
                            disabled={loading}
                          >
                            Edit
                          </button>
                          <button 
                            className={styles.iconButtonDelete}
                            onClick={() => this._deleteItem(item.Id!)}
                            title="Delete Item"
                            disabled={loading}
                          >
                            Delete
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

            </div>
          )}

        </div>
      </section>
    );
  }
}
