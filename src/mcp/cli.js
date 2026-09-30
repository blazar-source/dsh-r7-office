#!/usr/bin/env node

import { R7McpServer } from './server.js'

const server = new R7McpServer()
server.startStdio()
