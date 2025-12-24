"""
Модуль оптимизации маршрутов и расчета персонала
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import math
from typing import List, Dict, Tuple, Optional

class RouteOptimizer:
    """
    Оптимизатор маршрутов посещения точек
    """
    
    def __init__(self, data: pd.DataFrame):
        """
        Инициализация с данными из Excel
        """
        self.data = data
        self.points = []
        self.prepare_data()
        
    def prepare_data(self):
        """Подготовка данных для оптимизации"""
        for idx, row in self.data.iterrows():
            point = {
                'id': row.get('ID точки', idx),
                'name': row.get('Название точки', f'Точка_{idx}'),
                'address': row.get('Адрес', ''),
                'lat': float(row.get('Latitude', 0)),
                'lon': float(row.get('Longitude', 0)),
                'duration': float(row.get('Длительность посещения точки, минут', 0)),
                'frequency': int(row.get('Сколько раз надо посетить точку', 1)),
                'type': row.get('Название точки', 'Unknown')
            }
            self.points.append(point)
        print(f"Загружено {len(self.points)} точек")
    
    def calculate_distance(self, lat1: float, lon1: float, lat2: float, lon2: float) -> float:
        """
        Расчет расстояния между двумя точками (упрощенный, в км)
        Используем формулу гаверсинусов
        """
        R = 6371  # Радиус Земли в км
        
        lat1_rad = math.radians(lat1)
        lat2_rad = math.radians(lat2)
        delta_lat = math.radians(lat2 - lat1)
        delta_lon = math.radians(lon2 - lon1)
        
        a = (math.sin(delta_lat / 2) ** 2 + 
             math.cos(lat1_rad) * math.cos(lat2_rad) * 
             math.sin(delta_lon / 2) ** 2)
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        
        return R * c
    
    def estimate_travel_time(self, distance_km: float, avg_speed_kmh: float = 30) -> float:
        """
        Оценка времени в пути в минутах
        """
        return (distance_km / avg_speed_kmh) * 60
    
    def cluster_points_by_time(self, points: List[Dict], max_hours: float) -> List[List[Dict]]:
        """
        Кластеризация точек по времени (упрощенный алгоритм)
        """
        if not points:
            return []
            
        clusters = []
        current_cluster = []
        current_time = 0
        
        # Сортируем точки по близости к центру
        center_lat = np.mean([p['lat'] for p in points])
        center_lon = np.mean([p['lon'] for p in points])
        
        points_sorted = sorted(points, 
            key=lambda p: self.calculate_distance(p['lat'], p['lon'], center_lat, center_lon))
        
        max_minutes = max_hours * 60
        
        for point in points_sorted:
            point_time = point['duration']
            
            # Если точка не влезает в текущий кластер, начинаем новый
            if current_time + point_time > max_minutes * 0.8:  # 80% заполнения
                if current_cluster:
                    clusters.append(current_cluster)
                current_cluster = [point]
                current_time = point_time
            else:
                current_cluster.append(point)
                current_time += point_time
        
        if current_cluster:
            clusters.append(current_cluster)
            
        print(f"Создано {len(clusters)} кластеров")
        return clusters
    
    def solve_tsp(self, points: List[Dict]) -> List[Dict]:
        """
        Решение задачи коммивояжера (упрощенный жадный алгоритм)
        """
        if len(points) <= 1:
            return points
            
        # Начинаем с первой точки
        route = [points[0]]
        unvisited = points[1:]
        
        while unvisited:
            last_point = route[-1]
            
            # Находим ближайшую непосещенную точку
            nearest_idx = min(range(len(unvisited)),
                key=lambda i: self.calculate_distance(
                    last_point['lat'], last_point['lon'],
                    unvisited[i]['lat'], unvisited[i]['lon']
                ))
            
            route.append(unvisited[nearest_idx])
            unvisited.pop(nearest_idx)
        
        return route
    
    def distribute_visits_by_quarter(self, start_date: str = None) -> Dict[str, List[Dict]]:
        """
        Распределение посещений по дням квартала
        """
        if start_date is None:
            start_date = datetime.now().strftime('%Y-%m-%d')
        
        start = datetime.strptime(start_date, '%Y-%m-%d')
        quarter_days = []
        
        # Создаем список рабочих дней на квартал (90 дней)
        for i in range(90):
            current_date = start + timedelta(days=i)
            # Пропускаем выходные (суббота=5, воскресенье=6)
            if current_date.weekday() < 5:
                quarter_days.append(current_date.strftime('%Y-%m-%d'))
        
        print(f"Квартал: {len(quarter_days)} рабочих дней")
        
        # Собираем все посещения
        all_visits = []
        for point in self.points:
            for _ in range(point['frequency']):
                all_visits.append(point.copy())
        
        # Если частоты посещений = 0, добавляем каждую точку по 1 разу
        if not all_visits:
            all_visits = [p.copy() for p in self.points]
        
        print(f"Всего посещений: {len(all_visits)}")
        
        # Равномерно распределяем по дням
        visits_by_day = {}
        days_count = len(quarter_days)
        
        if days_count == 0:
            return {}
            
        for i, visit in enumerate(all_visits):
            day_index = i % days_count
            day = quarter_days[day_index]
            
            if day not in visits_by_day:
                visits_by_day[day] = []
            
            visits_by_day[day].append(visit)
        
        return visits_by_day
    
    def optimize_routes_for_day(self, day_points: List[Dict], max_hours: float, 
                               employee_id: int = 1) -> List[Dict]:
        """
        Оптимизация маршрутов для одного дня
        """
        if not day_points:
            return []
            
        # Кластеризуем точки по времени
        clusters = self.cluster_points_by_time(day_points, max_hours)
        
        routes = []
        route_id = 1
        
        for cluster in clusters:
            if not cluster:
                continue
                
            # Оптимизируем порядок посещения в кластере
            optimized_route = self.solve_tsp(cluster)
            
            # Расчет общего времени маршрута
            total_duration = sum(p['duration'] for p in optimized_route)
            
            # Расчет времени на дорогу
            travel_time = 0
            for i in range(len(optimized_route) - 1):
                dist = self.calculate_distance(
                    optimized_route[i]['lat'], optimized_route[i]['lon'],
                    optimized_route[i+1]['lat'], optimized_route[i+1]['lon']
                )
                travel_time += self.estimate_travel_time(dist)
            
            total_time = total_duration + travel_time
            
            route = {
                'route_id': f'R{employee_id}_{route_id}',
                'employee_id': f'Сотрудник_{employee_id}',
                'date': '',  # будет установлено позже
                'day_of_week': '',  # будет установлено позже
                'points': optimized_route,
                'total_points': len(optimized_route),
                'service_time_min': total_duration,
                'travel_time_min': travel_time,
                'total_time_min': total_time,
                'total_time_hours': round(total_time / 60, 2)
            }
            
            routes.append(route)
            route_id += 1
        
        return routes
    
    def assign_routes_to_employees(self, all_routes: List[Dict], 
                                  max_days_per_week: int) -> Tuple[List[Dict], int]:
        """
        Распределение маршрутов по сотрудникам
        """
        if not all_routes:
            return [], 0
            
        # Группируем маршруты по дням
        routes_by_day = {}
        for route in all_routes:
            day = route['date']
            if day not in routes_by_day:
                routes_by_day[day] = []
            routes_by_day[day].append(route)
        
        # Сортируем дни
        sorted_days = sorted(routes_by_day.keys())
        
        employees = []
        employee_counter = 1
        
        # Для каждого дня распределяем маршруты
        for day in sorted_days:
            day_routes = routes_by_day[day]
            
            # Пытаемся назначить существующим сотрудникам
            assigned = False
            for emp in employees:
                # Проверяем, не превышает ли сотрудник лимит дней
                emp_days = set([r['date'] for r in emp['routes']])
                if len(emp_days) < max_days_per_week and day not in emp_days:
                    # Назначаем все маршруты дня одному сотруднику
                    for route in day_routes:
                        route['employee_id'] = f'Сотрудник_{emp["id"]}'
                        emp['routes'].append(route)
                    assigned = True
                    break
            
            # Если не нашли подходящего сотрудника, создаем нового
            if not assigned:
                new_emp = {
                    'id': employee_counter,
                    'name': f'Сотрудник_{employee_counter}',
                    'routes': []
                }
                
                for route in day_routes:
                    route['employee_id'] = new_emp['name']
                    new_emp['routes'].append(route)
                
                employees.append(new_emp)
                employee_counter += 1
        
        # Преобразуем в плоский список маршрутов
        final_routes = []
        for emp in employees:
            final_routes.extend(emp['routes'])
        
        return final_routes, len(employees)
    
    def optimize(self, max_hours_per_day: float = 8, 
                max_days_per_week: int = 5,
                start_date: str = None) -> Dict:
        """
        Основной метод оптимизации
        """
        print(f"Начало оптимизации с параметрами: max_hours={max_hours_per_day}, max_days={max_days_per_week}")
        
        if start_date is None:
            start_date = datetime.now().strftime('%Y-%m-%d')
        
        # 1. Распределяем посещения по дням
        visits_by_day = self.distribute_visits_by_quarter(start_date)
        if not visits_by_day:
            return {
                'success': False,
                'error': 'Нет рабочих дней для планирования'
            }
        
        print(f"Распределено по {len(visits_by_day)} дней")
        
        # 2. Оптимизируем маршруты для каждого дня
        all_routes = []
        day_counter = 1
        
        for day, day_points in visits_by_day.items():
            if not day_points:
                continue
                
            print(f"Обработка дня {day_counter}/{len(visits_by_day)}: {len(day_points)} точек")
            
            # Создаем маршруты для дня (пока без привязки к сотрудникам)
            day_routes = self.optimize_routes_for_day(day_points, max_hours_per_day, 1)
            
            # Добавляем дату к маршрутам
            day_date = datetime.strptime(day, '%Y-%m-%d')
            for route in day_routes:
                route['date'] = day
                route['day_of_week'] = day_date.strftime('%A')
            
            all_routes.extend(day_routes)
            day_counter += 1
        
        print(f"Создано {len(all_routes)} маршрутов")
        
        if not all_routes:
            return {
                'success': False,
                'error': 'Не удалось создать маршруты'
            }
        
        # 3. Распределяем маршруты по сотрудникам
        final_routes, num_employees = self.assign_routes_to_employees(
            all_routes, max_days_per_week
        )
        
        print(f"Требуется сотрудников: {num_employees}")
        
        # 4. Подготавливаем результат
        result = {
            'success': True,
            'parameters': {
                'max_hours_per_day': max_hours_per_day,
                'max_days_per_week': max_days_per_week,
                'start_date': start_date
            },
            'summary': {
                'total_employees': num_employees,
                'total_routes': len(final_routes),
                'total_points': sum(len(r['points']) for r in final_routes),
                'avg_routes_per_employee': round(len(final_routes) / num_employees, 1) if num_employees > 0 else 0,
                'total_service_hours': round(sum(r['service_time_min'] for r in final_routes) / 60, 1),
                'total_travel_hours': round(sum(r['travel_time_min'] for r in final_routes) / 60, 1)
            },
            'routes': final_routes,
            'employees': [
                {
                    'id': i+1,
                    'name': f'Сотрудник_{i+1}',
                    'total_routes': len([r for r in final_routes if r['employee_id'] == f'Сотрудник_{i+1}']),
                    'total_days': len(set([r['date'] for r in final_routes if r['employee_id'] == f'Сотрудник_{i+1}']))
                }
                for i in range(num_employees)
            ]
        }
        
        return result