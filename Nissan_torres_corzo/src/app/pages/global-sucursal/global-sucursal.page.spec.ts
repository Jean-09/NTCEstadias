import { ComponentFixture, TestBed } from '@angular/core/testing';
import { GlobalSucursalPage } from './global-sucursal.page';

describe('GlobalSucursalPage', () => {
  let component: GlobalSucursalPage;
  let fixture: ComponentFixture<GlobalSucursalPage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(GlobalSucursalPage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
