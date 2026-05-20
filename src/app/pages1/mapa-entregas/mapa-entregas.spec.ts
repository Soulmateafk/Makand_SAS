import { ComponentFixture, TestBed } from '@angular/core/testing';

import { MapaEntregas } from './mapa-entregas';

describe('MapaEntregas', () => {
  let component: MapaEntregas;
  let fixture: ComponentFixture<MapaEntregas>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [MapaEntregas],
    }).compileComponents();

    fixture = TestBed.createComponent(MapaEntregas);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
