import { Component, Input, OnInit } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { AppServiceService } from '../app-service.service';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-addnewtask',
  templateUrl: './addnewtask.component.html',
  styleUrls: ['./addnewtask.component.css']
})
export class AddnewtaskComponent implements OnInit{
  @Input()userid!: number;
  public addForm !: FormGroup;
  constructor(private formBuilder: FormBuilder, private http: HttpClient, private router: Router, private service: AppServiceService) { }
  ngOnInit(): void {
    this.addForm = this.formBuilder.group({
      Name: ['', Validators.required],
      ProjectType: ['', Validators.required],
      Description: ['', Validators.required],
      Location: ['', Validators.required],
      SubDate: ['', Validators.required],
      SurveyHelpers: ['', Validators.required],
      model: ['', Validators.required],
      Reminder: ['', Validators.required]
    });
    
  }
  submit(){
    this.service.addProject(this.addForm.value).subscribe((res: any) => {
      if (res && res['status'] === 200) {
        console.log(res);
        alert("Project successfull added");
        this.addForm.reset();
        this.router.navigate(['head']);

      }else{
        console.log(res);
        alert("somthing wrong")

      }
    });

  }

}
