import { HttpClient } from '@angular/common/http';
import { Component, OnInit} from '@angular/core';
import { Router } from '@angular/router';
import { AppServiceService } from '../app-service.service';
import { ProjectList } from 'src/models/project-list.model';

@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css'],

})
export class TodoComponent implements OnInit{
  projectLists: any;
  projects: ProjectList[] = [];

  constructor(private http: HttpClient, private service: AppServiceService, private router: Router){}
  ngOnInit(){
    this.getProjectData();
  }
  getProjectData(): void {
    this.service.getProjectData(this.projects).subscribe(data=>{
      console.log(data);
      this.projectLists = data;
      this.router.navigate(['todo'])
      
    });
  };

}
 