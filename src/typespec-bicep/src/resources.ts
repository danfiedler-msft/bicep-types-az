// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  EmitContext,
  Enum,
  Model,
  ModelProperty,
  Namespace,
  Operation,
  Program,
  Type,
} from "@typespec/compiler";
import { ScopeType } from "bicep-types";
import {
  getArmResources,
  getArmResourceOperationList,
  resolveArmResources,
  ArmResourceDetails,
  ResolvedResource,
  isSingletonResource,
  getSingletonResourceKey,
} from "@azure-tools/typespec-azure-resource-manager";
import { getAllHttpServices, getHttpOperation, HttpOperation } from "@typespec/http";
import { BicepEmitterOptions } from "./lib.js";

/**
 * Describes the scope path segments for an ARM resource type.
 */
export interface ResourceDescriptor {
  namespace: string;
  typeSegments: string[];
  apiVersion: string;
  readableScopes: ScopeType;
  writableScopes: ScopeType;
  constantName?: string;
}

/**
 * A single ARM resource definition extracted from the TypeSpec program.
 */
export interface ResourceDefinition {
  descriptor: ResourceDescriptor;
  /** The TypeSpec Model representing the resource body (PUT request body). */
  putModel?: Model;
  /** The TypeSpec Model representing the resource response (GET response). */
  getModel?: Model;
  /** The name property on the resource, including any constraint decorators. */
  nameProperty?: ModelProperty;
}

/**
 * An ARM resource action (POST list* operation).
 */
export interface ResourceActionDefinition {
  actionName: string;
  descriptor: ResourceDescriptor;
  requestModel?: Model;
  responseModel?: Model;
}

/**
 * All resource definitions grouped by provider namespace and API version.
 */
export interface ProviderDefinition {
  namespace: string;
  apiVersion: string;
  resourcesByType: Record<string, ResourceDefinition[]>;
  resourceActions: ResourceActionDefinition[];
}

/**
 * Get the fully qualified resource type string (e.g. "Microsoft.Storage/storageAccounts").
 */
export function getFullyQualifiedType(descriptor: ResourceDescriptor): string {
  return [descriptor.namespace, ...descriptor.typeSegments].join("/");
}

/**
 * Extract all ARM provider definitions from a TypeSpec program.
 *
 * Uses the official @azure-tools/typespec-azure-resource-manager APIs
 * (getArmResources, resolveArmResources) to discover ARM resources,
 * their operations, and determine scopes from actual REST paths.
 */
export function getProviderDefinitions(
  context: EmitContext<BicepEmitterOptions>,
): ProviderDefinition[] {
  const program = context.program;
  const providers = new Map<string, ProviderDefinition>();

  // Use the ARM library's official resource discovery API
  const armResources = getArmResources(program);
  const resolvedProvider = resolveArmResources(program);
  const resolvedResources = resolvedProvider.resources ?? [];
  const customResourceRoutes = getCustomResourceRoutes(program);

  for (const armResource of armResources) {
    const namespace = armResource.armProviderNamespace;
    if (!namespace) continue;

    const model = armResource.typespecType;
    const apiVersion = getApiVersion(program, model);
    if (!apiVersion) continue;

    const key = `${namespace.toLowerCase()}/${apiVersion}`;
    if (!providers.has(key)) {
      providers.set(key, {
        namespace,
        apiVersion,
        resourcesByType: {},
        resourceActions: [],
      });
    }

    const provider = providers.get(key)!;

    // A single resource model can be exposed at multiple paths. Preserve every
    // resolved instance instead of selecting only the first one. Exclude
    // instances that merely share a route with the resource but whose actual
    // lifecycle operation returns a different model (e.g. a "backups" or
    // "commands" sub-path read that returns an unrelated response type) —
    // those aren't real instances of this resource.
    const resolvedMatches = resolvedResources.filter(
      (r) =>
        (r.type === model || r.resourceName === armResource.name) &&
        resolvedRepresentsModel(r, model),
    );

    // Determine scopes from operations and paths
    const { readableScopes, writableScopes } = getResourceScopesFromArm(
      program,
      armResource,
      resolvedProvider,
    );

    const routeMatches = customResourceRoutes.get(model) ?? [];
    if (
      readableScopes === ScopeType.None &&
      writableScopes === ScopeType.None &&
      routeMatches.length === 0
    ) {
      continue;
    }

    // Check for singleton resources (e.g. @singleton("default"))
    const isSingleton = isSingletonResource(program, model);
    const singletonKey = isSingleton ? getSingletonResourceKey(program, model) : undefined;

    const resourcePaths = resolvedMatches.length > 0
      ? resolvedMatches.map((resolved) => ({
          typeSegments: getResolvedTypeSegments(
            resolved.resourceType.types,
            armResource.collectionName,
          ),
          readableScopes,
          writableScopes,
        }))
      : routeMatches.length > 0
        ? routeMatches
        : [{
            typeSegments: armResource.collectionName
              ? [armResource.collectionName]
              : [],
            readableScopes,
            writableScopes,
          }];
    const seenResourcePaths = new Set<string>();

    for (const resourcePath of resourcePaths) {
      if (resourcePath.typeSegments.length === 0) continue;

      const expandedSegmentSets = expandParameterizedSegments(
        resourcePath.typeSegments,
        armResource,
      );

      for (const segments of expandedSegmentSets) {
        const fullyQualifiedType = `${namespace}/${segments.join("/")}`.toLowerCase();
        if (seenResourcePaths.has(fullyQualifiedType)) {
          continue;
        }
        seenResourcePaths.add(fullyQualifiedType);

        if (!provider.resourcesByType[fullyQualifiedType]) {
          provider.resourcesByType[fullyQualifiedType] = [];
        }

        const definition: ResourceDefinition = {
          descriptor: {
            namespace,
            typeSegments: segments,
            apiVersion,
            readableScopes: resourcePath.readableScopes,
            writableScopes: resourcePath.writableScopes,
            constantName: singletonKey,
          },
          putModel: model,
          getModel: model,
          nameProperty: getResourceNameProperty(model),
        };

        provider.resourcesByType[fullyQualifiedType].push(definition);

        // Discover actions for this resource
        discoverResourceActionsFromArm(armResource, provider, segments);
      }
    }
  }

  return [...providers.values()];
}

interface ResourceRoute {
  typeSegments: string[];
  readableScopes: ScopeType;
  writableScopes: ScopeType;
}

/**
 * Determines whether a resolved resource actually represents an instance of
 * `model`, as opposed to a synthetic entry created because some operation
 * (e.g. a read-styled action with extra path segments like a "backups" or
 * "commands" sub-path) shares the resource's route but returns a different
 * model. `resolveArmResources` groups operations purely by instance path, so
 * such operations get attached to the outer resource type even though their
 * request/response body is unrelated to it.
 */
function resolvedRepresentsModel(resolved: ResolvedResource, model: Model): boolean {
  const lifecycleOps = [
    ...(resolved.operations.lifecycle.createOrUpdate ?? []),
    ...(resolved.operations.lifecycle.read ?? []),
    ...(resolved.operations.lifecycle.update ?? []),
  ];
  if (lifecycleOps.length === 0) return false;

  return lifecycleOps.some((op) => getOperationResponseModel(op) === model);
}

function getResolvedTypeSegments(
  resolvedTypes: string[],
  collectionName: string | undefined,
): string[] {
  const typeSegments = [...resolvedTypes];
  if (
    collectionName &&
    typeSegments.at(-1)?.toLowerCase() !== collectionName.toLowerCase()
  ) {
    typeSegments.push(collectionName);
  }

  return typeSegments;
}

/**
 * Finds resource instance GET routes that are intentionally implemented as
 * custom HTTP operations instead of @armResourceOperations interfaces.
 */
function getCustomResourceRoutes(
  program: Program,
): Map<Model, ResourceRoute[]> {
  const routes = new Map<Model, ResourceRoute[]>();
  const [services, diagnostics] = getAllHttpServices(program);
  program.reportDiagnostics(diagnostics);

  for (const service of services) {
    for (const operation of service.operations) {
      if (operation.verb !== "get" || !isResourceInstancePath(operation.path)) {
        continue;
      }

      const typeSegments = getTypeSegmentsFromPath(operation.path);
      if (!typeSegments) continue;

      for (const responseModel of getResponseBodyModels(operation)) {
        const existing = routes.get(responseModel) ?? [];
        if (!existing.some((route) =>
          route.typeSegments.join("/").toLowerCase() ===
          typeSegments.join("/").toLowerCase()
        )) {
          existing.push({
            typeSegments,
            readableScopes: getScopeFromPath(operation.path),
            writableScopes: ScopeType.None,
          });
          routes.set(responseModel, existing);
        }
      }
    }
  }

  return routes;
}

function isResourceInstancePath(path: string): boolean {
  return path.split("/").filter(Boolean).at(-1)?.startsWith("{") === true;
}

function getTypeSegmentsFromPath(path: string): string[] | undefined {
  const segments = path.split("/").filter(Boolean);
  const providerIndex = segments.findIndex(
    (segment) => segment.toLowerCase() === "providers",
  );
  if (providerIndex < 0 || providerIndex + 2 >= segments.length) {
    return undefined;
  }

  const resourcePath = segments.slice(providerIndex + 2);
  const typeSegments = resourcePath.filter((_, index) => index % 2 === 0);
  return typeSegments.length > 0 ? typeSegments : undefined;
}

function* getResponseBodyModels(
  operation: HttpOperation,
): IterableIterator<Model> {
  for (const response of operation.responses) {
    for (const content of response.responses) {
      if (content.body?.type.kind === "Model") {
        yield content.body.type;
      }
    }
  }
}

/**
 * Expand parameterized type segments into concrete values.
 *
 * When a segment is a path parameter reference like "{recordType}", this
 * finds the corresponding enum (or union of string literals) from the
 * resource's operations and expands it into separate segment arrays.
 *
 * For example, if typeSegments is ["dnsZones", "{recordType}"] and the
 * RecordType enum has values [A, AAAA, CAA, ...], this returns
 * [["dnsZones", "A"], ["dnsZones", "AAAA"], ["dnsZones", "CAA"], ...].
 *
 * If no parameterized segments exist, returns the original array wrapped
 * in a single-element array.
 */
function expandParameterizedSegments(
  typeSegments: string[],
  armResource: ArmResourceDetails,
): string[][] {
  // Find parameterized segments (segments wrapped in {})
  const paramIndex = typeSegments.findIndex(
    (seg) => seg.startsWith("{") && seg.endsWith("}"),
  );
  if (paramIndex === -1) {
    return [typeSegments];
  }

  const paramName = typeSegments[paramIndex].slice(1, -1);
  const enumValues = resolvePathParameterEnum(paramName, armResource);
  if (!enumValues || enumValues.length === 0) {
    // Cannot resolve — keep the parameterized segment as-is
    return [typeSegments];
  }

  // Replace the parameterized segment with each concrete value
  const result: string[][] = [];
  for (const value of enumValues) {
    const expanded = [...typeSegments];
    expanded[paramIndex] = value;
    result.push(expanded);
  }
  return result;
}

/**
 * Resolve the concrete string values for a path parameter enum.
 *
 * Searches the resource's lifecycle and action operations for a @path
 * parameter matching the given name, then extracts enum member values.
 */
function resolvePathParameterEnum(
  paramName: string,
  armResource: ArmResourceDetails,
): string[] | undefined {
  // Collect all operations to search through
  const ops = armResource.operations;
  const allOps: { httpOperation: { parameters: { parameters: { param: { name: string; type: Type } }[] } } }[] = [];

  if (ops.lifecycle.read) allOps.push(ops.lifecycle.read);
  if (ops.lifecycle.createOrUpdate) allOps.push(ops.lifecycle.createOrUpdate);
  if (ops.lifecycle.update) allOps.push(ops.lifecycle.update);
  if (ops.lifecycle.delete) allOps.push(ops.lifecycle.delete);
  for (const action of Object.values(ops.actions)) {
    allOps.push(action);
  }
  for (const list of Object.values(ops.lists)) {
    allOps.push(list);
  }

  for (const op of allOps) {
    for (const httpParam of op.httpOperation.parameters.parameters) {
      if (httpParam.param.name === paramName) {
        return extractEnumValues(httpParam.param.type);
      }
    }
  }

  return undefined;
}

/**
 * Extract string values from an Enum or Union type.
 */
function extractEnumValues(type: Type): string[] | undefined {
  if (type.kind === "Enum") {
    const enumType = type as Enum;
    const values: string[] = [];
    for (const member of enumType.members.values()) {
      const value = typeof member.value === "string" ? member.value : member.name;
      values.push(value);
    }
    return values.length > 0 ? values : undefined;
  }

  if (type.kind === "Union") {
    const values: string[] = [];
    for (const variant of type.variants.values()) {
      if (variant.type.kind === "String") {
        values.push(variant.type.value);
      } else if (variant.type.kind === "EnumMember") {
        const value = typeof variant.type.value === "string" ? variant.type.value : variant.type.name;
        values.push(value);
      }
    }
    return values.length > 0 ? values : undefined;
  }

  return undefined;
}

/**
 * Determine readable and writable scopes for a resource by examining its
 * resolved operations and their HTTP paths.
 */
function getResourceScopesFromArm(
  _program: Program,
  armResource: ArmResourceDetails,
  resolvedProvider: ReturnType<typeof resolveArmResources>,
): { readableScopes: ScopeType; writableScopes: ScopeType } {
  const declaredOperations = [...getArmResourceOperationList(
    _program,
    armResource.typespecType,
  )];
  if (armResource.name === "TestType1") {
    console.error("scope-debug", declaredOperations.map((operation) => [operation.kind, operation.path, getHttpOperation(_program, operation.operation)[0].path]));
  }

  // Try to find resolved resource details with path information
  const resolvedResources = resolvedProvider.resources ?? [];
  const matched = resolvedResources.filter(
    (r) =>
      r.type === armResource.typespecType ||
      r.resourceName === armResource.name,
  );

  if (matched.length > 0) {
    let readableScopes = ScopeType.None;
    let writableScopes = ScopeType.None;

    for (const resolved of matched) {
      const scope = getScopeFromPath(resolved.resourceInstancePath);
      const ops = resolved.operations;

      // Check for read operations
      if (ops.lifecycle?.read) {
        readableScopes |= scope;
      }

      // Check for create/update operations (write)
      if (ops.lifecycle?.createOrUpdate || ops.lifecycle?.update) {
        writableScopes |= scope;
      }
    }

    // Resolved metadata can omit lifecycle operations registered through legacy
    // helpers, so merge it with the authoritative ARM operation metadata below.
    for (const operation of declaredOperations) {
      const scope = getArmOperationScope(_program, operation);
      if (operation.kind === "read") {
        readableScopes |= scope;
      } else if (operation.kind === "createOrUpdate" || operation.kind === "update") {
        writableScopes |= scope;
      }
    }

    if (armResource.operations.lifecycle.createOrUpdate || armResource.operations.lifecycle.update) {
      writableScopes |= getDefaultScopeFromKind(armResource.kind);
    }

    if (armResource.name === "TestType1") console.error("scope-result", readableScopes, writableScopes);
    return { readableScopes, writableScopes };
  }

  // Fall back: use lifecycle operations from ArmResourceDetails
  const hasRead = declaredOperations.some((operation) => operation.kind === "read");
  const hasWrite = declaredOperations.some(
    (operation) => operation.kind === "createOrUpdate" || operation.kind === "update",
  );

  // Determine scope from the resource kind and any available path
  let scope = getDefaultScopeFromKind(armResource.kind);

  // Refine scope from read operation path if available
  const firstLifecycleOperation = declaredOperations.find(
    (operation) => operation.kind === "read" || operation.kind === "createOrUpdate" || operation.kind === "update",
  );
  if (firstLifecycleOperation) {
    scope = getArmOperationScope(_program, firstLifecycleOperation);
  }

  return {
    readableScopes: hasRead ? scope : ScopeType.None,
    writableScopes: hasWrite ? scope : ScopeType.None,
  };
}

function getArmOperationScope(
  program: Program,
  operation: {
    path?: string;
    operation: Operation;
    resourceKind?: "legacy" | "legacy-extension";
  },
): ScopeType {
  if (operation.path) {
    return getScopeFromPath(operation.path);
  }

  const [httpOperation, diagnostics] = getHttpOperation(
    program,
    operation.operation,
  );
  program.reportDiagnostics(diagnostics);
  return httpOperation.path
    ? getScopeFromPath(httpOperation.path)
    : getDefaultScopeFromKind(operation.resourceKind ?? "Proxy");
}

/**
 * Determine scope from an HTTP path by analyzing the path prefix.
 */
function getScopeFromPath(path: string): ScopeType {
  // Management group: /providers/Microsoft.Management/managementGroups/{mgId}/providers/...
  if (path.match(/\/providers\/Microsoft\.Management\/managementGroups\//i)) {
    return ScopeType.ManagementGroup;
  }

  // Resource group: /subscriptions/{sub}/resourceGroups/{rg}/providers/...
  if (path.match(/\/subscriptions\/[^/]+\/resourceGroups\//i)) {
    return ScopeType.ResourceGroup;
  }

  // Subscription: /subscriptions/{sub}/providers/...
  if (path.match(/\/subscriptions\/[^/]+\/providers\//i)) {
    return ScopeType.Subscription;
  }

  // Tenant: /providers/... (no subscription or RG prefix)
  if (path.match(/^\/providers\//i)) {
    return ScopeType.Tenant;
  }

  // Extension scope: {resourceUri}/providers/...
  if (path.match(/\{[^}]+\}\/providers\//i)) {
    return ScopeType.Extension;
  }

  return ScopeType.ResourceGroup;
}

/**
 * Get default scope based on ARM resource kind.
 */
function getDefaultScopeFromKind(kind: string): ScopeType {
  switch (kind) {
    case "Extension":
      return ScopeType.Extension;
    case "Tracked":
      return ScopeType.ResourceGroup;
    case "Proxy":
      return ScopeType.ResourceGroup;
    default:
      return ScopeType.ResourceGroup;
  }
}

/**
 * Discover POST actions from ARM resource metadata.
 * Only POST operations (like listKeys) become resource function types.
 * GET operations (even if marked @action) are list operations, not functions.
 */
function discoverResourceActionsFromArm(
  armResource: ArmResourceDetails,
  provider: ProviderDefinition,
  typeSegments: string[],
): void {
  const actions = armResource.operations.actions;
  for (const [actionName, action] of Object.entries(actions)) {
    // Only include POST actions — GET @action operations are list operations
    const httpVerb = action.httpOperation.verb;
    if (httpVerb !== "post") {
      continue;
    }

    const responseModel = getOperationResponseModel(action);
    const requestModel = getOperationRequestModel(action);

    provider.resourceActions.push({
      actionName,
      descriptor: {
        namespace: provider.namespace,
        typeSegments,
        apiVersion: provider.apiVersion,
        readableScopes: ScopeType.ResourceGroup,
        writableScopes: ScopeType.ResourceGroup,
      },
      requestModel,
      responseModel,
    });
  }
}

/**
 * Get the API version for a resource model.
 */
function getApiVersion(
  _program: Program,
  model: Model,
): string | undefined {
  const ns = model.namespace;
  if (!ns) return undefined;

  // Look for a "Versions" enum in the namespace or parent namespaces
  let current: Namespace | undefined = ns;
  while (current) {
    const versionsEnum = current.enums.get("Versions");
    if (versionsEnum) {
      const members = [...versionsEnum.members.values()];
      if (members.length > 0) {
        const lastMember = members[members.length - 1];
        const value = lastMember.value ?? lastMember.name;
        return typeof value === "string" ? value : String(value);
      }
    }
    current = current.namespace;
  }

  return undefined;
}

/**
 * Get the 'name' property on a resource model.
 */
function getResourceNameProperty(
  model: Model,
): ModelProperty | undefined {
  return model.properties.get("name");
}

/**
 * Get the response model from an ARM resource operation.
 */
function getOperationResponseModel(
  action: { httpOperation: { responses: readonly { type?: Type }[] } },
): Model | undefined {
  for (const response of action.httpOperation.responses) {
    if (response.type && response.type.kind === "Model") {
      return response.type;
    }
  }
  return undefined;
}

/**
 * Get the request body model from an ARM resource operation.
 */
function getOperationRequestModel(
  action: { httpOperation: { parameters: { body?: { type?: Type } } } },
): Model | undefined {
  const body = action.httpOperation.parameters.body;
  if (body?.type?.kind === "Model") {
    return body.type;
  }
  return undefined;
}
