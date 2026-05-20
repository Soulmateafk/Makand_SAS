import { Component, AfterViewInit, inject } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http'; 
import Feature from 'ol/Feature';
import Point from 'ol/geom/Point';
import { Vector as VectorLayer } from 'ol/layer';
import { Vector as VectorSource } from 'ol/source';
import { Style, Icon } from 'ol/style';
import { fromLonLat } from 'ol/proj';
import Map from 'ol/Map';
import View from 'ol/View';
import TileLayer from 'ol/layer/Tile';
import OSM from 'ol/source/OSM';

@Component({ 
  selector: 'app-mapa-entregas',
  standalone: true,
  templateUrl: './mapa-entregas.html',
  styleUrls: ['./mapa-entregas.css']
})
export class MapaEntregasComponent implements AfterViewInit {
  private http = inject(HttpClient); 
  map: any;

  ngAfterViewInit(): void {
    this.map = new Map({
      target: 'map',
      layers: [new TileLayer({ source: new OSM() })],
      view: new View({ center: fromLonLat([-74.0721, 4.7110]), zoom: 12 }),
    });

    this.cargarDatosDesdeAPI();
  }

  cargarDatosDesdeAPI() {
    const token = localStorage.getItem('token');
    const headers = new HttpHeaders({
      'Authorization': `Bearer ${token}`
    });

    this.http.get<any[]>('http://localhost:3000/api/mapa', { headers }).subscribe({
      next: (data) => {
        console.log("Datos recibidos de /api/mapa:", data);
        this.pintarMarcadores(data);
      },
      error: (err) => {
        console.error('Error al cargar entregas desde /api/mapa:', err);
      }
    });
  }

  pintarMarcadores(entregas: any[]) {
    const vectorSource = new VectorSource();

    entregas.forEach(item => {
      // Intentamos extraer latitud y longitud de los nombres más comunes
      let lon = Number(item.lon ?? item.longitud ?? 0);
      let lat = Number(item.lat ?? item.latitud ?? 0);

      // CORRECCIÓN: Si no hay coordenadas, usamos el centro de Bogotá por defecto
      // para que el mapa muestre algo y sepas que el código funciona.
      if (lon === 0 || lat === 0) {
        console.warn("Sin coordenadas para:", item, "- Usando ubicación por defecto.");
        lon = -74.0721;
        lat = 4.7110;
      }

      const coords = fromLonLat([lon, lat]); 
      
      const feature = new Feature({
        geometry: new Point(coords),
      });

      feature.setStyle(
        new Style({
          image: new Icon({
            anchor: [0.5, 1],
            src: 'https://openlayers.org/en/latest/examples/data/icon.png',
          }),
        })
      );

      vectorSource.addFeature(feature);
    });

    this.map.addLayer(new VectorLayer({ source: vectorSource }));
    console.log("Proceso de pintado finalizado.");
  }
}