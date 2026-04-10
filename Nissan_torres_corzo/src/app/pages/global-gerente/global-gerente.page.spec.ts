import { ComponentFixture, TestBed } from '@angular/core/testing';
import { GlobalGerentePage } from './global-gerente.page';

describe('GlobalGerentePage', () => {
  let component: GlobalGerentePage;
  let fixture: ComponentFixture<GlobalGerentePage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(GlobalGerentePage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
