import { HttpClient } from '@angular/common/http';
import { Component, OnInit, ViewChild } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { switchAll } from 'rxjs';
import { AppServiceService } from '../app-service.service';
import { HeadComponent } from '../head/head.component';


@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.css']
})
export class LoginComponent implements OnInit {

  selecteduserId! : number;
  
  public loginForm!: FormGroup
  userName: any;
  Name: any;
  
  constructor(private formBuilder: FormBuilder, private http: HttpClient, private router: Router, private service: AppServiceService) {}

  ngOnInit(): void {
    this.loginForm = this.formBuilder.group({
      email: ['', Validators.required],
      password: ['', Validators.required]
    })
    
  }
  login() {
    this.service.getOneUserData(this.loginForm.value).subscribe((res:any) => {
      if(res && res['status'] === 200){
        const userid = res.userId;
        this.selecteduserId = userid;
        const Name = res.userName;
        this.userName = Name;
        console.log(userid, Name);
        alert("Successfull Login");
        this.loginForm.reset();
        this.router.navigate(['head']);
        //localStorage.setItem(userid, Name);
        return;
      }else{
        console.log(res);}
        alert("Wrong Details")
      }
    );     
  }

}
function showErrorAlert() {
  throw new Error('Function not implemented.');
}

