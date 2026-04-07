import { ComponentFixture, TestBed } from '@angular/core/testing';
import { APVPage } from './apv.page';

describe('APVPage', () => {
  let component: APVPage;
  let fixture: ComponentFixture<APVPage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(APVPage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
