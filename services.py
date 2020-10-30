import wmi
import sys
import time

class WindowsDAO(object):
	'''
	Acesso ao Servidor Windows
	'''

	def __init__(self):

		self.ip = self.getCredentials('resources.db.windows.params.host')
		self.user = self.getCredentials('resources.db.windows.params.user')
		self.passw = self.getCredentials('resources.db.windows.params.password')

		self.connection = self.connect()

		self.timeOut = 60

	def connect(self):
		'''
		Cria conexão com o servidor
		'''
		
		try:
			connect = wmi.WMI(self.ip, user = self.user, password = self.passw)

		except wmi.x_wmi:
			connect = None

		return connect

	def getCredentials(self, credential):
		'''
		Obtém credenciais
		'''

		# necessário alterar conforme necessidade
		file = open("../application/application.ini", "r")

		for line in file.readlines():
			if line.split(' = ')[0] == credential:
				file.close()
				return (line.split(' = ')[1].replace('\n', '').replace('"', ''))

		file.close()

	def getService(self, serviceName):
		'''
		Obtém serviço
		'''

		sql = f"SELECT * FROM Win32_Service where DisplayName = '{serviceName}' OR Name = '{serviceName}' "

		return self.connection.query(sql)

	def getAllServices(self):
		'''
		Obtém todos os serviços de uma lista
		'''

		services = ['a', 'b', 'c']

		json = {}

		for service in services:

			json['service'] = self.getService(service)[0]

		return json

	def startService(self, service):
		'''
		Inicia serviço
		'''

		if service.State == 'Stopped':

			service.StartService()

		else:

			raise Exception(f"erro ao iniciar o serviço {service.DisplayName}.")

	def stopService(self, service):
		'''
		Para serviço
		'''

		if service.State == 'Running':

			service.StopService()

		else:

			raise Exception(f"erro ao parar o serviço {service.DisplayName}.")

	def restartService(self, service):
		'''
		Reinicia serviço
		'''

		if service.State == 'Running':

			self.stopService(service)
			internalTimeout = 0

			while service.State != 'Stopped' and internalTimeout <= self.timeOut:
				time.sleep(1)
				internalTimeout += 1
				service = self.getService(service.DisplayName)[0]

			if internalTimeout >= self.timeOut:
				raise Exception(f"sessão expirada. {internalTimeout} segundos.")

			self.startService(service)

if __name__ == "__main__":

	windows = WindowsDAO()
	service = windows.getService(sys.argv[1])[0]
	action = sys.argv[2]

	if action == 'restart':

		try:

			windows.restartService(service)
			service = self.getService(service.DisplayName)[0]
			print(service.State)

		except Exception as e:

			print(f'não foi possível reiniciar o serviço {service}, {e}')
			service = windows.getService(service.DisplayName)[0]
			print(service.State)

	elif action == 'start':

		try:

			netsystem.startService(service)
			service = windows.getService(service.DisplayName)[0]
			print(service.State)

		except Exception as e:

			print(f'não foi possível iniciar o serviço {service}')
			service = self.getService(service.DisplayName)[0]
			print(service.State)

	elif action == 'stop':

		try:

			netsystem.stopService(service)
			service = windows.getService(service.DisplayName)[0]
			print(service.State)

		except Exception as e:

			print(f'não foi possível parar o serviço {service}')
			service = windows.getService(service.DisplayName)[0]
			print(service.State)

	else:

		print('comando inválido')
		print(service)
