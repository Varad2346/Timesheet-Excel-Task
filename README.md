# Demo

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 18.2.21.

## Development server

Run `ng serve` for a dev server. Navigate to `http://localhost:4200/`. The application will automatically reload if you change any of the source files.

## Code scaffolding

Run `ng generate component component-name` to generate a new component. You can also use `ng generate directive|pipe|service|class|guard|interface|enum|module`.

## Build

Run `ng build` to build the project. The build artifacts will be stored in the `dist/` directory.

## Running unit tests

Run `ng test` to execute the unit tests via [Karma](https://karma-runner.github.io).

## Running end-to-end tests

Run `ng e2e` to execute the end-to-end tests via a platform of your choice. To use this command, you need to first add a package that implements end-to-end testing capabilities.

## Further help

To get more help on the Angular CLI use `ng help` or go check out the [Angular CLI Overview and Command Reference](https://angular.dev/tools/cli) page.



  // let minDate=new Date(data[0].DATE);
    // let maxDate=new Date(data[0].DATE);
    // data.forEach(entry=>{
      
    //   const currDate=new Date(entry.DATE);
    //   if(currDate.getTime()>=minDate.getTime()){
    //     maxDate=currDate
    //   }else if(currDate.getTime()<=minDate.getTime()){
    //     minDate=currDate
    //   }
    // })
    // const diff=maxDate.getTime()-minDate.getTime();

    // const totalCalendarDays = Math.round(diff / (1000 * 60 * 60 * 24)) + 1;    
  
    // console.log(totalCalendarDays)

    // this.data.map((data)=>{
    //     let customerName=data['Customer Name']
    //     if(customerNames.indexOf(customerName)==-1){
    //       if(cnt==1){
    //         customerNames+=`${cnt}.${customerName}`;
    //       }else{
    //         customerNames+=`, ${cnt}.${customerName}`;
    //       }
    //       cnt++;
    //     }

    //     const exist=transformedData.find((row)=>{
    //       if(row.DATE==data.Date){
    //         return row
    //       }else {
    //         return {}
    //       }
    //     })

    //     if(exist){
    //       exist={...exist,"TASK DESCRIPTION":exist['TASK DESCRIPTION'+data['TASK DESCRIPTION']]}
    //     }
    //     // console.log(exist,data['S.No.'])
    //     const entry= {
    //       "Sr.No":srNo++,
    //       "DATE":data.Date,
    //       "DAY":new Date(data.Date).toLocaleDateString('en-us',{weekday:'long'}),
    //       "Efforts(HOURS)":parseFloat(data.Hours)+(parseFloat(data.Minutes)/60),
    //       "TASK DESCRIPTION":`${data['Task Description']}`
    //     }
    //     transformedData.push(entry)
    //   })
    //   console.log(transformedData)


    // const aoa=[
    //   ["CUSTOMER NAME",customerNames,"PROJECT MANAGER","VARAD","CALENDAR DAYS",""],
    //   ["RESOURCE NAME","AJIT","APPROVER NAME","VEDANT","WEEKLY OFF/HOLIDAYS",""],
    //   ["FOR MONTH","","SUBMITTED BY","AJIT","WORKED DAYS",""],
    //   ["ROLE","FRONTEND","SUBMISSION DATE","","LEAVES TAKEN",""]
    // ]

  //   let holidays=0
  //   let workingDays=0
  //   const transformedData = this.data.map((data) => {
      
  //     const dateObj = new Date(data.Date);
  //     const formattedDate = `${dateObj.getDate()}-${this.monthNames[dateObj.getMonth()]}-${dateObj.getFullYear().toString().slice(-2)}`;
  //     const day = this.dayNames[dateObj.getDay()];
  //     const efforts = parseFloat((data.Hours || 0)) +(parseFloat((data.Minutes || 0) )/60);
  //     console.log(efforts)
  //     if(efforts==0) {
  //       holidays=holidays+1
  //     }else{
  //       workingDays++
  //     }
  //      return {
  //       "Sr.No": data['S.No.'],
  //       "DATE": formattedDate,
  //       "DAY": day,
  //       "EFFORTS (Hours)": efforts,
  //       "TASK DESCRIPTION": data['Task Description']
  //     };
  //   });

  //   const dateObj = new Date(this.data[0].Date);
  //   const calendarDays = new Date(dateObj.getFullYear(), dateObj.getMonth() + 1, 0).getDate();
  //   const aoa = [
  //     ['CUSTOMER NAME', this.data[0]['Customer Name'], 'PROJECT MANAGER', 'VARAAD', 'CALENDAR DAYS',calendarDays],
  //     ['RESOURCE NAME', 'AJIT', 'APPROVER', 'VARAD', 'WEEKLY OFF / HOLIDAYS', holidays],
  //     ['FOR MONTH', `${this.monthNames[dateObj.getMonth()]}-${dateObj.getFullYear()}`, 'SUBMITTED BY', 'AJIT', 'WORKED DAYS',workingDays ],
  //     ['ROLE', 'FRONT END', 'SUBMISSION DATE', '', 'LEAVES TAKEN', 0],
  //   ];
  //   console.log(aoa)
  //   aoa.unshift(['Credenca Data Solutions']); 


    // const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(aoa);

 

    // const wb: XLSX.WorkBook = XLSX.utils.book_new();
    // XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  // }