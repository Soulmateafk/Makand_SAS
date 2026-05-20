import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Combustible } from './combustible';

describe('Combustible', () => {
  let component: Combustible;
  let fixture: ComponentFixture<Combustible>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Combustible],
    }).compileComponents();

    fixture = TestBed.createComponent(Combustible);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
