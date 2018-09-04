import datetime
import calendar

import requests


class Invoice(object):
    """
    Clase que se encarga de obtener el total de facturación anual
    """
    # Endpoint a la api
    endpoint = ''
    # Payload para el endpoint
    payload = {}
    # Contador para el total de facturas
    total_invoices = 0
    # Contador para las peticiones realizadas
    total_request = 0
    # Año actual
    current_year = None

    def __init__(self, endpoint, payload):
        """
        Constructor
        """
        self.endpoint = endpoint
        self.payload = payload
        self.current_year = datetime.datetime.now().year

    def _make_request(self):
        """
        Ejecuta la petición al servidor
        """
        return requests.get(self.endpoint, params=self.payload)

    def _get_total_days_month(self, month):
        """
        Obtiene el total de días de un mes
        """
        return calendar.monthrange(self.current_year, month)[1]

    def _get_total_weeks_month(self, month):
        """
        Obtiene el total de semanas de un mes
        """
        last_day = self._get_total_days_month(month)
        date = datetime.datetime(self.current_year, month, last_day)
        return (date.day-1)//7+1

    def get_total_invoices(self):
        """
        Obtengo el total de facturas del año actual
        """
        # Si no pudo obtener el total anual
        if not self._get_total_invoices_year():
            # Recorro por los 12 meses
            for month in range(1, 13):
                # Proceso el mes
                print("Procesando el mes: " + str(month))
                self._get_total_invoices_month(month)

        # Retorno el total de facturas y las peticiones realizadas
        return {
            'total_invoices': str(self.total_invoices),
            'total_request': str(self.total_request)
        }

    def _get_total_invoices_year(self):
        """
        Obtengo el total de facturas por año
        """
        start = datetime.datetime(self.current_year, 1, 1)
        finish = datetime.datetime(self.current_year, 12, 31)

        start_format = start.strftime("%Y-%m-%d")
        finish_format = finish.strftime("%Y-%m-%d")
        # Cargo el período en los parámetros
        self.payload['start'] = start_format
        self.payload['finish'] = finish_format

        print("Procesando por año")
        print("Procesando el período " + start_format + " / " + finish_format)

        # Ejecuto el endpoint
        r = self._make_request()
        # Incremento en 1 el contador de peticiones
        self.total_request += 1

        try:
            # Sumarizo el total de facturas
            self.total_invoices += int(r.text)
            return True
        except ValueError:
            return False

    def _get_total_invoices_month(self, month):
        """
        Obtiene el total de facturas por mes
        """
        # Obtengo los parámetro start and finish
        start = datetime.datetime(self.current_year, month, 1)
        finish = datetime.datetime(
            self.current_year, month, self._get_total_days_month(month)
        )

        start_format = start.strftime("%Y-%m-%d")
        finish_format = finish.strftime("%Y-%m-%d")
        # Cargo el período en los parámetros
        self.payload['start'] = start_format
        self.payload['finish'] = finish_format

        print("Procesando por mes")
        print("Procesando el período " + start_format + " / " + finish_format)

        # Ejecuto el endpoint
        r = self._make_request()
        # Incremento en 1 el contador de peticiones
        self.total_request += 1

        try:
            # Sumarizo el total de facturas
            self.total_invoices += int(r.text)
        except ValueError:
            self._get_total_invoices_week(month, 1, 1)

    def _get_total_invoices_week(self, month, week, first_day):
        """
        Obtiene el total de facturas por semana
        """
        start = datetime.datetime(self.current_year, month, first_day)
        is_last_week = False
        # Si es la última semana
        if week == self._get_total_weeks_month(month):
            is_last_week = True
            # Obtengo el último día del mes
            finish = datetime.datetime(
                self.current_year, month, self._get_total_days_month(month)
            )
        else:
            # Voy a la siguiente semana
            finish = datetime.datetime(
                self.current_year, month, first_day + 6
            )

        start = start.strftime("%Y-%m-%d")
        finish = finish.strftime("%Y-%m-%d")
        # Cargo el período en los parámetros
        self.payload['start'] = start
        self.payload['finish'] = finish

        print("Procesando por semana")
        print("Procesando el período " + start + " / " + finish)

        # Ejecuto el endpoint
        r = self._make_request()
        # Incremento en 1 el contador de peticiones
        self.total_request += 1

        try:
            # Sumarizo el total de facturas
            self.total_invoices += int(r.text)
            # Mientras no sea la última semana del mes
            if not is_last_week:
                self._get_total_invoices_week(month, week + 1, first_day + 7)
        except ValueError:
            self._get_total_invoices_day(month, 1)

    def _get_total_invoices_day(self, month, day):
        """
        Obtiene el total de facturas por día
        """
        start = datetime.datetime(self.current_year, month, day)
        finish = datetime.datetime(self.current_year, month, day + 1)
        start = start.strftime("%Y-%m-%d")
        finish = finish.strftime("%Y-%m-%d")
        # Cargo el período en los parámetros
        self.payload['start'] = start
        self.payload['finish'] = finish

        print("Procesando pro día")
        print("Procesando el período " + start + " / " + finish)

        # Ejecuto el endpoint
        r = self._make_request()
        # Incremento en 1 el contador de peticiones
        self.total_request += 1

        try:
            # Sumarizo el total de facturas
            self.total_invoices += int(r.text)
            self._get_total_invoices_day(month, day + 1)
        except ValueError:
            pass

