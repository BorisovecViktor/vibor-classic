class EmployeesWP {
  constructor() {
    this.employeesList;
    this.employeesBodyId = "emplBody";
    this.employeesListId = "empList";
    this.employeesHeadersId = "empHeaders";
    this.listTitle = "Employees";
    this.endPoint = `/_api/lists/getbytitle('${this.listTitle}')/items`;
  }

  async sandBox() {
    const url = `${_spPageContextInfo.webAbsoluteUrl}/`;

    //Не разобрался почему я не могу вытащить с инпута его текущее value, у меня всегда во всех полях оно null
    //возможно это особенности sharepoint?
    //Так же не понял почему я не могу создавать тег form, я бы мог просто обработать форму вместо того что б
    //создавать переменную на каждый инпут что б достать value

    // let title = document.getElementById('title');
    // let alias = document.getElementById('alias').value;
    // let position = document.getElementById('position').value;
    // let gender = document.getElementById('gender').value;
    // let office = document.getElementById('office').value;

    // const newEmployee = {
    //   Title: title,
    //   Alias: alias,
    //   Position: position,
    //   Gender: gender,
    //   OfficeId: office,
    // }

    //получаем список сотрудников
    try {
      const query = `${url}_api/web/lists/getbytitle('Employees')/items`;
      const result = await this.getItems(query);
      this.employeesList = result.d.results;
    } catch (err) {
      console.log(err);
    }

    //добавляем нового сотрудника
    const btn = document.getElementById('form__btn');
    btn.addEventListener('click', () => {
      try {
        const createdata = this.createData(url);
        console.log(createdata);
      } catch (err) {
        console.log(err);
      }
    })

    //рендерим страницу
    try {
      this.renderHTML();
    } catch (err) {
      console.log(err);
    }
  }

  getItems(query) {
    return $.ajax({
      url: query,
      method: "GET",
      contentType: "application/json;odata=verbose",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });
  }

  //формируем сотрудника
  async createData(webUrl) {
    const query = `${webUrl}_api/web/lists/getbytitle('Employees')/items`;
    const requestDigest = await this.getRequestDigest(webUrl);
    const listItemType = await this.getListItemType(
      webUrl,
      "Employees"
    );
    const changes = {
      Title: 'Oleg Shevchenko',
      Alias: 'olshe',
      Position: 'Manager',
      Gender: 'Male',
      OfficeId: 3,
    };
    const objType = {
      __metadata: {
        type: listItemType.d.ListItemEntityTypeFullName,
      },
    };
    const objData = JSON.stringify(Object.assign(objType, changes));

    return $.ajax({
      url: query,
      type: "POST",
      data: objData,
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest":
          requestDigest.d.GetContextWebInformation.FormDigestValue,
        "X-HTTP-Method": "POST",
      },
    });
  }

  getRequestDigest(webUrl) {
    return $.ajax({
      url: webUrl + "_api/contextinfo",
      method: "POST",
      headers: {
        Accept: "application/json; odata=verbose",
      },
    });
  }

  getListItemType(url, listTitle) {
    const query =
      url +
      "_api/Web/Lists/getbytitle('" +
      listTitle +
      "')/ListItemEntityTypeFullName";
    return this.getItems(query);
  }

  renderHTML() {
    try {
      //печатаем заголовки
      const headers = ['Name', 'Alias', 'Position', 'Sertificate', 'Gender', 'Office'];
      let employeeHeaders = "";

      headers.map((header) => {
        const headerItem = new EmployeeHeader(header);
        employeeHeaders += headerItem.getHTML();
      });

      document.getElementById(
        this.employeesHeadersId
      ).innerHTML += employeeHeaders;

      //печатаем список сотрудников
      let employeeItems = "";

      this.employeesList.map((employee) => {
        const employeeItem = new EmployeeItem(employee);
        employeeItems += employeeItem.getHTML();
      });

      document.getElementById(
        this.employeesListId
      ).innerHTML += employeeItems;
    } catch (error) {
      console.log("dprtWrapper", "No data provided");
    }
  }
}

//формируем заголовок
class EmployeeHeader {
  constructor(employeeHeader) {
    this.header = employeeHeader;
  }

  getHTML() {
    let employeeHeaderTemplate = `      
      <div>
        <span>${this.header}</span>
      </div>        
    `;
    return employeeHeaderTemplate;
  }
}

//формируем поля сотрудника
class EmployeeItem {
  constructor(employeeItem) {
    this.title = employeeItem.Title;
    this.alias = employeeItem.Alias;
    this.position = employeeItem.Position;
    this.sertificate = employeeItem.Sertificate;
    this.gender = employeeItem.Gender;
    this.office = employeeItem.OfficeId;
  }

  getHTML() {
    let employeeItemTemplate = `      
      <div class="employees__item">
        <div class="employees__field">
          <span>${this.title}</span>
        </div>
        <div class="employees__field">
          <span>${this.alias}</span>
        </div>
        <div class="employees__field">
          <span>${this.position}</span>
        </div>
        <div class="employees__field">
          <span>${this.sertificate ? 'yes' : 'no'}</span>
        </div>
        <div class="employees__field">
          <span>${this.gender}</span>
        </div>
        <div class="employees__field">
          <span>${this.office}</span>
        </div>
      </div>
    `;
    return employeeItemTemplate;
  }
}

SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {
  const emplWP = new EmployeesWP();
  emplWP.sandBox();
});
