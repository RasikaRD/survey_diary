import { HttpClient } from '@angular/common/http';
import { Component } from '@angular/core';
import { Router } from '@angular/router';
import { ProjectReport } from 'src/models/project-list.model';
import { AppServiceService } from '../app-service.service';
import * as XLSX from 'xlsx';
import { Element } from '@angular/compiler';

@Component({
  selector: 'app-workdone',
  templateUrl: './workdone.component.html',
  styleUrls: ['./workdone.component.css']
})

export class WorkdoneComponent {

  projectReports: any;
  projects: ProjectReport[] = [];
  fileName = 'Report.xlsx';
  lastElement: any;
  trElement: any;

  constructor(private http: HttpClient, private service: AppServiceService, private router: Router) { }

  ngOnInit() {
    this.getProjectReport();
  }

  getProjectReport(): void {
    this.service.getProjectReport(this.projects).subscribe(data => {
      this.projectReports = data;
    });
  };

  onDown(index: number): void {

    // remove download tabel row
    const trId = 'down-' + index;
    this.trElement = document.getElementById(trId);
    if(this.trElement != null){
      this.trElement.remove();
    }

    const tableId = 'project-report-' + index;
    let element = document.getElementById(tableId);

    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(element);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'report1');
    XLSX.writeFile(wb, this.fileName);
  };
}


