import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import {HttpClientModule} from '@angular/common/http';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { HeadComponent } from './head/head.component';
import { LoginComponent } from './login/login.component';
import { SingupComponent } from './singup/singup.component';
import { AuthModule } from './auth/auth/auth.module';
import { TodoComponent } from './todo/todo.component';
import { WorkdoneComponent } from './workdone/workdone.component';
import { AddnewtaskComponent } from './addnewtask/addnewtask.component';


@NgModule({
  declarations: [
    AppComponent,
    HeadComponent,
    LoginComponent,
    SingupComponent,
    TodoComponent,
    WorkdoneComponent,
    AddnewtaskComponent,

  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    FormsModule,
    ReactiveFormsModule,
    HttpClientModule,
    AuthModule,
    
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
