#!/usr/bin/python
# -*- coding: UTF-8 -*-

import paramiko
import sys
import time

class TotvsDAO(object):

	def __init__(self):

		self.host = 'host'
		self.port = 'port'
		self.user = 'user'
		self.passw = 'password'

		self.ssh = self.connect()

	def connect(self):

		ssh = paramiko.SSHClient()
		ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
		ssh.connect(self.host, self.port, self.user, self.passw)

		return ssh

	def disconnect(self):

		self.ssh.close()

	def getServers(self):

		serverList = []

		command = "cd /usr/local/etc; cat desconecta_bancos-prod.pf;"
		stdin, stdout, stderr = self.ssh.exec_command(command, get_pty = True)
		response = stdout.readlines()

		for row in response:

			row = row.rstrip('\r\n')
			column = row.split(' ')
			serverList.append(column[1])

		return serverList

	def getUserConnections(self, servers, user):

		userList = []

		for server in servers:

			command = f"cd /usr/local/bin; /opt/dlc116/bin/proshut {server} -C list;"

			stdin, stdout, stderr = self.ssh.exec_command(command, get_pty = True)
			response = stdout.readlines()

			#    0        1         2          3        4         5         6          7         8            9            10
			#  'usr'  | 'pid' | 'weekday' | 'month' | 'day' |   'hour'  | 'year' | 'user id' | 'Type' |     'tty'      | 'Limbo?'
			#  2454   | 4160  |    Tue    |   Nov   |   3   |  04:34:07 |  2020  |   jonas   |  REMC  | NDBPODT1260776 |   no

			for row in response:
				row = ' '.join(row.split()).rstrip('\r\n')
				column = row.split(' ')

				if column[7] == user:

					userList.append((column[0], column[7], server))

		return userList

	def disconnectUser(self, userToDc):

		serverList = self.getServers()
		userList = self.getUserConnections(serverList, userToDc)

		for user in userList:

			command = f"cd /usr/local/bin; /opt/dlc116/bin/proshut {user[2]} -C disconnect {user[0]}"

			stdin, stdout, stderr = self.ssh.exec_command(command, get_pty = True)
			response = stdout.readlines()
		
		self.disconnect()

if __name__ == "__main__":

	user = sys.argv[1]

	totvs = TotvsDAO()

	servers = totvs.getUserConnections(totvs.getServers(), user)

	totvs.disconnectUser(user)

	if len(servers) > 0:

		print('true')

	else:

		print('false')

