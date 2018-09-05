//
//  PersonService.swift
//  Proximity
//
//  Created by Fernando Romiti on 28/05/2018.
//  Copyright © 2018 Estimote, Inc. All rights reserved.
//

import Foundation
import Moya
import SwiftyJSON

struct Person {
    var apellido: String
    var nombres: String
    var mail: String
    var usuario: String
    var cdCliente: Int
    var cdPersona: Int
    
    init() {
        self.apellido = ""
        self.nombres = ""
        self.mail = ""
        self.usuario = ""
        self.cdCliente = 0
        self.cdPersona = 0
    }
    
    init(apellido: String, nombres: String, mail: String, usuario: String, cdCliente: Int, cdPersona: Int) {
        self.apellido = apellido
        self.nombres = nombres
        self.mail = mail
        self.usuario = usuario
        self.cdCliente = cdCliente
        self.cdPersona = cdPersona
    }
}

enum PersonService {
    case register(token: String, person: Person)
}

extension PersonService: BaseService {
    typealias ParsedModel = Person
    
    var baseURL: URL {
        let str = ServiceConstants.BaseURL.Evan
        return URL(string: str)!
    }
    
    var path: String {
        switch self {
        case .register:
            return "Users/NewPerson"
        }
    }
    
    var method: Moya.Method {
        switch self {
        case .register:
            return .post
        }
    }
    
    var headers: [String : String]? {
        switch self {
        case .register(let token, _):
            return ["Authorization": "Bearer \(token)"]
        }
    }
    
    var parameters: [String: Any]? {
        switch self {
        case .register(_, let person):
            return ["NmApellido": person.apellido,
                    "NmNombres": person.nombres,
                    "NmMail": person.mail,
                    "CdUsuario": person.usuario,
                    "CdClient": person.cdCliente,
                    "CdPersona": person.cdPersona]
        }
    }
    
    var parameterEncoding: ParameterEncoding {
        switch self {
        case .register:
            return URLEncoding.default
        }
    }
    
    var sampleData: Data {
        switch self {
        case .register:
            return FileReader.readFileFrom(filename: "person")
        }
    }
    
    var task: Task {
        switch self {
        case .register:
            return .requestPlain
        }
    }
    
    var parseJSON: (JSON) -> (Person?) {
        switch self {
        case .register:
            return { json in
                if let ok = json["Ok"].bool, ok {
                    let apellido = json["Data"]["NmApellido"].string ?? ""
                    let nombres = json["Data"]["NmNombres"].string ?? ""
                    let mail = json["Data"]["NmMail"].string ?? ""
                    let usuario = json["Data"]["CdUsuario"].string ?? ""
                    let cliente = json["Data"]["CdClient"].int ?? 1
                    let persona = json["Data"]["CdPersona"].int ?? 0
                    return Person(apellido: apellido, nombres: nombres, mail: mail, usuario: usuario, cdCliente: cliente, cdPersona: persona)
                }
                return nil
            }
        }
    }
    
}
