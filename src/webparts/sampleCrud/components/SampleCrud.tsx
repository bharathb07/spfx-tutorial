import * as React from 'react';
import { ISampleCrudProps } from './ISampleCrudProps';
import { getSP, SPFI } from '../../../pnpjsConfig';
import { PrimaryButton, DetailsList, SelectionMode, IconButton, Dialog, DialogType, DialogFooter, TextField, DefaultButton, DatePicker, IDatePickerStrings } from 'office-ui-fabric-react';


const datePickerStrings: IDatePickerStrings = {
  // Customize strings as per requirements
  months: [
    'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'
  ],
  shortMonths: [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
  ],
  days: [
    'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'
  ],
  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker'
};

// type definitions
interface IEmployeeDetailsRes {
  Id: number;
  Title: number;
  Employee_name: string;
  Role: string;
  Country: string;
  DOJ: string;

}

interface ISampleProps {
  Id: number;
  employee_id: number;
  employee_name: string;
  role: string;
  country: string;
  DOJ: string;
}

interface IFormState {
  employee_id: string;
  employee_name: string;
  role: string;
  country: string;
  DOJ: Date | undefined;
}

const SampleCrud = (props: ISampleCrudProps): React.ReactElement => {
  const _sp: SPFI = getSP(props.spcontext);
  const [reload, setReload] = React.useState<boolean>(false);
  const [data, setData] = React.useState<Array<ISampleProps>>([]);
  const [isAddHidden, setIsAddHidden] = React.useState<boolean>(true);
  const [currentId, setCurrentId] = React.useState<number | any>();
  const [isEditHidden, setIsEditHidden] = React.useState<boolean>(true);
  const [deleteId, setDeleteId] = React.useState<number | any>();
  const [isDeleteHidden, setIsDeleteHidden] = React.useState<boolean>(true);

  const [formState, setFormState] = React.useState<IFormState>({
    employee_id: '',
    employee_name: '',
    role: '',
    country: '',
    DOJ: undefined
  });


  const getListItems = async () => {
    try {
      const listItems = await _sp.web.lists.getByTitle('EmployeeDetails').items();
      setData(listItems.map((each: IEmployeeDetailsRes) => ({
        Id: each.Id,
        employee_id: each.Title,
        employee_name: each.Employee_name,
        role: each.Role,
        country: each.Country,
        DOJ: each.DOJ
      })))
    } catch (e) {
      console.error(e);
    }
  };

  React.useEffect(() => {
    const fetchData = async () => {
      await getListItems();
    };
    fetchData().catch((error) => console.error('Error fetching data:', error));
  }, [reload]);

  const resetFormState = () => {
    setFormState({
      employee_id: '',
      employee_name: '',
      role: '',
      country: '',
      DOJ: undefined,
    });
  }


  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = event.target;
    setFormState((prevState) => ({
      ...prevState,
      [name]: value,
    }));
  };
  const handleDateChange = (date: Date | null | undefined) => {
    setFormState((prevState) => ({
      ...prevState,
      DOJ: date || undefined,
    }));
  };


  const addNewListItem = async () => {
    const list = _sp.web.lists.getByTitle("EmployeeDetails");
    try {
      await list.items.add({
        Title: formState.employee_id,
        Employee_name: formState.employee_name,
        Role: formState.role,
        Country: formState.country,
        DOJ: formState.DOJ ? formState.DOJ.toISOString() : null,
      });
      setIsAddHidden(true);
      setReload(!reload);
      resetFormState();
      console.log('List item added');
    } catch (e) {
      console.log(e);
    }
  }

  const openEditDialog = (id: number) => {
    setCurrentId(id)
    setIsEditHidden(false);
    const employeeData = data?.find((each: ISampleProps) => each.Id === id);
    if (employeeData) {
      setFormState({
        employee_id: employeeData.employee_id.toString(),
        employee_name: employeeData.employee_name,
        role: employeeData.role,
        country: employeeData.country,
        DOJ: employeeData.DOJ ? new Date(employeeData.DOJ) : undefined,
      });
    }
  };


  const editListItem = async () => {
    if (currentId === null) return;
    const list = _sp.web.lists.getByTitle("EmployeeDetails");
    try {
      await list.items.getById(currentId).update({
        Employee_name: formState.employee_name,
        Role: formState.role,
        Country: formState.country,
        DOJ: formState.DOJ?.toISOString(), // Ensure this field exists and is correctly formatted
      });

      setIsEditHidden(true);
      setReload(!reload);
      resetFormState();
      console.log('List item edited');
    } catch (e) {
      console.log(e);
    }
  };

  const openDeleteDialog = (id: number) => {
    setDeleteId(id);
    setIsDeleteHidden(false);
  }

  const deleteListItem = async () => {
    const list = _sp.web.lists.getByTitle("EmployeeDetails");
    try {
      await list.items.getById(deleteId).delete();
      setIsDeleteHidden(true);
      setReload(!reload);
      setDeleteId(null);
      console.log('List item deleted');
    } catch (e) {
      console.log(e);
    }
  }

  return (
    <>
      <h1>Add and update Employee Details </h1>
      <div className='quoteBox'>
        <h2>Employee Details</h2>
        <div className='quoteContainer'>
          <DetailsList
            items={data || []}
            columns={[
              {
                key: 'employeeIdColumn',
                name: 'Employee_id',
                fieldName: 'employee_id',
                minWidth: 10,
                isResizable: true,
                onRender: (item: ISampleProps) => <div>{item.employee_id}</div>,
              },
              {
                key: 'employeeNameColumn',
                name: 'Employee_name',
                fieldName: 'employee_name',
                minWidth: 50,
                isResizable: true,
                onRender: (item: ISampleProps) => <div>{item.employee_name}</div>,
              },
              {
                key: 'employeeRoleColumn',
                name: 'Role',
                fieldName: 'role',
                minWidth: 50,
                isResizable: true,
                onRender: (item: ISampleProps) => <div>{item.role}</div>,
              },
              {
                key: 'employeeCountryColumn',
                name: 'Country',
                fieldName: 'country',
                minWidth: 50,
                isResizable: true,
                onRender: (item: ISampleProps) => <div>{item.country}</div>,
              },
              {
                key: 'employeeDOJColumn',
                name: 'DOJ',
                fieldName: 'DOJ',
                minWidth: 50,
                isResizable: true,
                onRender: (item: ISampleProps) => <div>{item.DOJ ? new Date(item.DOJ).toLocaleDateString() : ''}</div>,
              },
              {
                key: 'actionsColumn',
                name: 'Actions',
                minWidth: 100,
                isResizable: true,
                onRender: (item: ISampleProps) => (
                  <div>
                    <IconButton
                      iconProps={{ iconName: 'Edit' }}
                      onClick={() => openEditDialog(item.Id)}
                      title="Edit"
                      ariaLabel="Edit"
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      onClick={() => openDeleteDialog(item.Id)} // handles the delete list item functionality
                      title="Delete"
                      ariaLabel="Delete"
                    />
                  </div>
                ),
              },
            ]}
            selectionMode={SelectionMode.none}
          />

          {/* Edit Employee details popup and form ... */}
          <Dialog
            hidden={isEditHidden}
            onDismiss={() => setIsEditHidden(true)}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Edit Employee Details',
            }}
          >
            <div>
              <TextField
                label="Employee_id"
                name="employee_id"
                value={formState.employee_id}
                onChange={handleInputChange}
              />
              <TextField
                label="Employee_name"
                name="employee_name"
                value={formState.employee_name}
                onChange={handleInputChange}
              />
              <TextField
                label="Role"
                name="role"
                value={formState.role}
                onChange={handleInputChange}
              />
              <TextField
                label="Country"
                name="country"
                value={formState.country}
                onChange={handleInputChange}
              />
              <DatePicker
                label="Date of Joining"
                strings={datePickerStrings}
                value={formState.DOJ}
                onSelectDate={handleDateChange}
                isRequired={false}
                placeholder="Select a date..."
              />
            </div>
            <DialogFooter>
              <PrimaryButton text="Submit" onClick={() => editListItem()} />
              <DefaultButton text="Cancel" onClick={() => {
                setIsEditHidden(true)
                resetFormState()
              }} />
            </DialogFooter>
          </Dialog>


          {/* Delete the Employee details popup... */}
          <Dialog
            hidden={isDeleteHidden}
            onDismiss={() => setIsDeleteHidden(true)}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Delete Employee Details',
            }}
          >
            <div>
              <p>Are sure want to delete this Details?</p>
            </div>
            <DialogFooter>
              <PrimaryButton text="Submit" onClick={() => deleteListItem()} />
              <DefaultButton text="Cancel" onClick={() => {
                setIsDeleteHidden(true)
                setDeleteId(null)
              }} />
            </DialogFooter>
          </Dialog>
        </div>


        {/* Add employee details popup and form... */}
        <div>
          <PrimaryButton text='Add Employee' onClick={() => setIsAddHidden(false)} />
        </div>
        <Dialog
          hidden={isAddHidden}
          onDismiss={() => setIsAddHidden(true)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Add Employee Details',
          }}
        >
          <div>
            <TextField
              label="Employee_id"
              name="employee_id"
              value={formState.employee_id}
              onChange={handleInputChange}
            />
            <TextField
              label="Employee_name"
              name="employee_name"
              value={formState.employee_name}
              onChange={handleInputChange}
            />
            <TextField
              label="Role"
              name="role"
              value={formState.role}
              onChange={handleInputChange}
            />
            <TextField
              label="Country"
              name="country"
              value={formState.country}
              onChange={handleInputChange}
            />
            <DatePicker
              label="Date of Joining"
              strings={datePickerStrings}
              value={formState.DOJ}
              onSelectDate={handleDateChange}
              isRequired={false}
              placeholder="Select a date..."
            />
          </div>
          <DialogFooter>
            <PrimaryButton text="Submit" onClick={() => addNewListItem()} />
            <DefaultButton text="Cancel" onClick={() => {
              setIsAddHidden(true)
              resetFormState()
            }
            } />
          </DialogFooter>
        </Dialog>
      </div>

    </>
  )
}
export default SampleCrud;