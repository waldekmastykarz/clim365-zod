import fs from 'fs';
import parse from 'yargs-parser';
import { ZodTypeAny, z } from 'zod';

// we need to use a function or we lose the schema type information
function alias<T extends ZodTypeAny>(alias: string, type: T) {
  type._def.alias = alias;
  return type;
}

// #region parsing schema to get the necessary info for yargs-parser and loading commands
function parseEffect(def: z.ZodEffectsDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  return def.schema._def;
}

function parseIntersection(def: z.ZodIntersectionDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (def.left._def.typeName !== z.ZodFirstPartyTypeKind.ZodAny) {
    return def.left._def;
  }

  if (def.right._def.typeName !== z.ZodFirstPartyTypeKind.ZodAny) {
    return def.right._def;
  }

  return undefined;
}

function parseObject(def: z.ZodObjectDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  const properties = def.shape();
  for (const key in properties) {
    const property = properties[key];

    const option: OptionInfo = {
      name: key,
      alias: property._def.alias,
      required: true,
      type: 'string'
    };

    parseDef(property._def, options, option);
    options.push(option);
  }

  return;
}

function parseString(def: z.ZodStringDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'string';
  }
  return;
}

function parseNumber(def: z.ZodNumberDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'number';
  }
  return;
}

function parseBoolean(def: z.ZodBooleanDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'boolean';
  }
  return;
}

function parseOptional(def: z.ZodOptionalDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.required = false;
  }

  return def.innerType._def;
}

function parseDefault(def: z.ZodDefaultDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  return def.innerType._def;
}

function parseEnum(def: z.ZodEnumDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'string';
    currentOption.autocomplete = def.values;
  }

  return;
}

function parseNativeEnum(def: z.ZodNativeEnumDef, options: OptionInfo[], currentOption?: OptionInfo): z.ZodTypeDef | undefined {
  if (currentOption) {
    currentOption.type = 'string';
    currentOption.autocomplete = Object.getOwnPropertyNames(def.values);
  }

  return;
}

function getParseFn(typeName: z.ZodFirstPartyTypeKind) {
  switch (typeName) {
    case z.ZodFirstPartyTypeKind.ZodEffects:
      return parseEffect;
    case z.ZodFirstPartyTypeKind.ZodObject:
      return parseObject;
    case z.ZodFirstPartyTypeKind.ZodOptional:
      return parseOptional;
    case z.ZodFirstPartyTypeKind.ZodString:
      return parseString;
    case z.ZodFirstPartyTypeKind.ZodNumber:
      return parseNumber;
    case z.ZodFirstPartyTypeKind.ZodBoolean:
      return parseBoolean;
    case z.ZodFirstPartyTypeKind.ZodEnum:
      return parseEnum;
    case z.ZodFirstPartyTypeKind.ZodNativeEnum:
      return parseNativeEnum;
    case z.ZodFirstPartyTypeKind.ZodDefault:
      return parseDefault;
    case z.ZodFirstPartyTypeKind.ZodIntersection:
      return parseIntersection;
    default:
      return;
  }
}

function parseDef(def: z.ZodTypeDef, options: OptionInfo[], currentOption?: OptionInfo) {
  let parsedDef: z.ZodTypeDef | undefined = def;

  do {
    const parse = getParseFn((parsedDef as any).typeName);
    if (!parse) {
      break;
    }

    parsedDef = parse(parsedDef as any, options, currentOption);
    if (!parsedDef) {
      break;
    }

  } while (parsedDef);
}

function schemaToOptions(schema: z.ZodSchema<any>): OptionInfo[] {
  const options: OptionInfo[] = [];
  parseDef(schema._def, options);
  return options;
}
// #endregion

// interface that represents information that goes into allCommands.json
// and allCommandsFull.json
interface OptionInfo {
  name: string;
  alias?: string;
  required: boolean;
  autocomplete?: string[];
  type: 'string' | 'boolean' | 'number';
}

// #region global options that apply to all commands
const OutputType = z.enum(['csv', 'json', 'md', 'text', 'none']);
type OutputType = z.infer<typeof OutputType>;

const globalOptions = z.object({
  query: z.string().optional(),
  output: OutputType.optional(),
  debug: z.boolean().default(false),
  verbose: z.boolean().default(false)
});

function getGlobalRefinedSchema(schema: typeof globalOptions) {
  return schema
    .refine(options => !options.debug || !options.verbose, {
      message: 'Specify debug or verbose, but not both'
    });
}
// #endregion

// login command example
export enum CloudType {
  Public,
  USGov,
  USGovHigh,
  USGovDoD,
  China
}

// login command options extending global options
const commandOptions = globalOptions
  .extend({
    authType: alias('t', z.enum(['certificate', 'deviceCode', 'password', 'identity', 'browser', 'secret']).optional().default('deviceCode')),
    cloud: z.nativeEnum(CloudType).optional().default(CloudType.Public),
    userName: alias('u', z.string().optional()),
    password: alias('p', z.string().optional()),
    certificateFile: alias('c', z.string().optional()
      .refine(filePath => !filePath || fs.existsSync(filePath), filePath => ({
        message: `Certificate file ${filePath} does not exist`
      }))),
    certificateBase64Encoded: z.string().optional(),
    thumbprint: z.string().optional(),
    appId: z.string().optional(),
    tenant: z.string().optional(),
    secret: alias('s', z.string().optional()),
    dummyNumber: z.number().optional(),
    dummyBoolean: z.boolean().optional()
  })
  // don't allow unknown properties; default for all commands
  .strict();
  // if we want to allow unknown properties, remove the .strict() call and
  // uncomment the following line
  // .and(z.any());

function getRefinedSchema(schema: typeof commandOptions) {
  return (getGlobalRefinedSchema(schema as any) as unknown as typeof commandOptions)
    .refine(options => options.authType !== 'password' || options.userName, {
      message: 'Username is required when using password authentication'
    })
    .refine(options => options.authType !== 'password' || options.password, {
      message: 'Password is required when using password authentication'
    })
    .refine(options => options.authType !== 'certificate' || !(options.certificateFile && options.certificateBase64Encoded), {
      message: 'Specify either certificateFile or certificateBase64Encoded, but not both.'
    })
    .refine(options => options.authType !== 'certificate' || options.certificateFile || options.certificateBase64Encoded, {
      message: 'Specify either certificateFile or certificateBase64Encoded'
    })
    .refine(options => options.authType !== 'secret' || options.secret, {
      message: 'Secret is required when using secret authentication'
    });
}

// sample command line arguments
const emptyArgs: string[] = [];
const debugArgs = ['--debug'];
const userNamePasswordArgs = ["--authType", "password", "--userName", "user@contoso.com", "--password", "pass@word1"];
const pemCertArgs = ["--authType", "certificate", "--certificateFile", "/Users/user/dev/localhost.pem"];
const certThumbprintArgs = ["--authType", "certificate", "--certificateFile", "/Users/user/dev/localhost.pem", "", "--thumbprint", "47C4885736C624E90491F32B98855AA8A7562AF"];
const pfxArgs = ["--authType", "certificate", "--certificateFile", "/Users/user/dev/localhost.pfx", "--password", "pass@word1"];
const certBase64 = ["--authType", "certificate", "--certificateBase64Encoded", "MIII2QIBAzCCCJ8GCSqGSIb3DQEHAaCCCJAEgeX1N5AgIIAA==", "--thumbprint", "D0C9B442DE249F55D10CDA1A2418952DC7D407A3"];
const identityArgs = ["--authType", "identity"];
const userAssignedIdentityArgs = ["--authType", "identity", "--userName", "ac9fbed5-804c-4362-a369-21a4ec51109e"];
const invalidAuthType = ["--authType", "invalid"];
const invalidOutput = ["--output", "invalid"];
const output = ["--output", "json"];
const unknownOptions = ["--Title", "new list item"];
const userNamePasswordAliasArgs = ["-t", "password", "-u", "user@contoso.com", "-p", "pass@word1"];

// translate the schema to allCommands.json
const optionsInfo = schemaToOptions(commandOptions);

// parse the command line arguments
const argv = parse(userNamePasswordAliasArgs, {
  // build a list of key-value pairs for aliases
  alias: optionsInfo.reduce((aliases: { [key: string]: string }, option) => {
    if (option.alias) {
      aliases[option.name] = option.alias;
    }
    return aliases;
  }, {}),
  // make it explicit how to parse each option value
  boolean: optionsInfo.filter(option => option.type === 'boolean').map(option => option.name),
  number: optionsInfo.filter(option => option.type === 'number').map(option => option.name),
  string: optionsInfo.filter(option => option.type === 'string').map(option => option.name),
  configuration: {
    "parse-numbers": false,
    "strip-aliased": true,
    "strip-dashed": true
  }
});
// we're not using positional args in CLI for M365
delete (argv as any)._;

// validate against command's schema
const result = getRefinedSchema(commandOptions).safeParse(argv);
if (result.success) {
  console.log(result.data);
}
else {
  console.error(`Error in property ${result.error.errors[0].path}: ${result.error.errors[0].message}`);
}

// example typesafe enum from zod
switch (argv.output) {
  case OutputType.enum.csv:
    console.log('csv');
    break;
  case OutputType.enum.json:
    console.log('json');
    break;
  case OutputType.enum.md:
    console.log('md');
    break;
  case OutputType.enum.text:
    console.log('text');
    break;
  case OutputType.enum.none:
    console.log('none');
    break;
}