import { HttpClient } from '@angular/common/http';
import { Component } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { AppServiceService } from '../app-service.service';


@Component({
  selector: 'app-singup',
  templateUrl: './singup.component.html',
  styleUrls: ['./singup.component.css']
})
export class SingupComponent {
  public singupForm !: FormGroup;
  constructor(private formBuilder: FormBuilder, private http: HttpClient, private router: Router, private service: AppServiceService) { }
  ngOnInit(): void {
    this.singupForm = this.formBuilder.group({
      fullname: ['', Validators.required],
      email: ['', Validators.required],
      mobile: ['', Validators.required],
      password: ['', Validators.required]
    })
  }

  singup() {

    this.service.addUserData(this.singupForm.value).subscribe((res: any) => {
      if (res && res['status'] === 200) {
        console.log(res);
        alert("Singup is successfull");
        this.singupForm.reset();
        this.router.navigate(['login']);

      }else{
        console.log(res);
        alert("Email already registered")

      }
    });


  }

}



