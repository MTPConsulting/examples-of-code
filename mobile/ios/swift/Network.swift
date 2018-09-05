//
//  Network.swift
//  Proximity
//
//  Created by Fernando Romiti on 24/05/2018.
//  Copyright © 2018 Estimote, Inc. All rights reserved.
//

import Foundation
import Moya
import SwiftyJSON

struct Network<T: BaseService> {
    typealias EndpointClosure = (T) -> Endpoint
    
    static func request(
        service: T,
        provider: MoyaProvider<T>? = nil,
        success successCallback: @escaping (T.ParsedModel?) -> (),
        error errorCallback: @escaping (MoyaError) -> (),
        failure failureCallback: @escaping (MoyaError) -> ()
        ) {
        
        let endpointClosure = { (target: T) -> Endpoint in
            return MoyaProvider.defaultEndpointMapping(for: service)
        }
        
        let requestProvider = provider ?? MoyaProvider<T>(endpointClosure: endpointClosure,
                                                          manager: DefaultAlamofireManager.sharedManager)
        
        requestProvider.request(service, callbackQueue: DispatchQueue.global()){ result in
            switch result {
            case let .success(response):
                do {
                    // see if the response has 200-299 status code
                    let _ = try response.filterSuccessfulStatusCodes()
                    // use SwiftyJSON to get JSON from response
                    // TODO: Catch JSON errors
                    let json = try JSON(data: response.data)
                    let parsedResult = service.parseJSON(json)
                    successCallback(parsedResult)
                } catch let error as MoyaError {
                    errorCallback(error)
                } catch {
                    print("Unknown error: \(error)")
                }
            case let .failure(error):
                failureCallback(error)
            }
        }
    }
}
