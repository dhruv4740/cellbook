import { hashPassword } from '../../src/renderer/src/model/protection'
process.stdout.write(JSON.stringify(hashPassword(process.argv[2])))
