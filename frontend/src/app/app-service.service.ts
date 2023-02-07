import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { map } from 'rxjs/operators';
import { ProjectList } from 'src/models/project-list.model';

@Injectable({
  providedIn: 'root'
})
export class AppServiceService {

  constructor(private http: HttpClient) { }


  addUserData(payload:object){
    return this.http.post(`http://localhost:3000/singing`, payload).pipe(map(res=>{
      return res;
    }))
  }
  getOneUserData(payload:object){
    return this.http.post(`http://localhost:3000/login`, payload).pipe(map(res=>{
      return res;
    }))
  }

  addProject(payload: Object){
    return this.http.post(`http://localhost:3000/addnew`, payload).pipe(map(res=>{
      return res;
    }))
  }

  getProjectData(payload: Object){
    return this.http.get<ProjectList>(`http://localhost:3000/projects`, payload).pipe(map(res=>{
      return res;
    }))

  }

}
