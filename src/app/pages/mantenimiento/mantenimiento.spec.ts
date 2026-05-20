import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Mantenimiento } from './mantenimiento';

describe('Mantenimiento', () => {
  let component: Mantenimiento;
  let fixture: ComponentFixture<Mantenimiento>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Mantenimiento],
    }).compileComponents();

    fixture = TestBed.createComponent(Mantenimiento);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
